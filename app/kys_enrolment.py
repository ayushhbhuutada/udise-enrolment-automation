# -*- coding: utf-8 -*-
"""
UDISE+ Know Your School  ->  Grade-wise enrolment collector

For every school in the .xlsx list next to start.bat (sheet INPUT_SHEET) it:
  1. opens https://kys.udiseplus.gov.in/#/advancesearch (School Name search)
  2. fills State + School Name, then WAITS FOR YOU to type the captcha (Search is auto-clicked)
  3. picks the result whose UDISE code matches the list's School Code / UDISE column
  4. opens the Report Card and reads
     "Enrolment by grade in the current academic session (by Age in completed years)"
  5. writes one column per grade into output/<list name> - Enrolment.xlsx (a copy of the list)

Re-running continues where it stopped. To redo a school, clear its Status cell
in the output workbook and run again.
"""
import re
import sys
import json
import time
import datetime
import subprocess
from copy import copy
from pathlib import Path
from urllib.parse import urljoin

try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    from playwright.sync_api import sync_playwright, Error as PWError
except ImportError as e:
    print("Missing library:", e)
    print("Run start.bat (it installs everything automatically).")
    sys.exit(1)

# ----------------------------- SETTINGS ------------------------------------
INPUT_SHEET = "Focused Data"       # sheet to process (first sheet is used if this one is missing)
SAVE_EVERY_N_SCHOOLS = 1           # save the output workbook after this many schools (name-search mode)

# BULK MODE: ONE captcha lists every school of the district (Advance Search); then every report
# card is opened directly - no more captchas. Schools not in that list fall back to name search.
BULK_MODE = True
BULK_STATE = "MAHARASHTRA"
BULK_DISTRICT = "PUNE"
BULK_DELAY_SECONDS = 1.0           # polite pause between report cards
BULK_SAVE_EVERY = 10               # in bulk mode save after this many schools (and when stopping)
ID_CACHE_DAYS = 7                  # reuse the district list for this many days (no captcha on resume)
DEFAULT_YEAR_SEG = "13"            # report-card address year part, used only if it could not be learnt
DEFAULT_STATE = "MAHARASHTRA"      # used when the State column is empty / missing
MAX_CANDIDATES_TO_FETCH = 10       # if several schools match, fetch data for up to this many
RETRY_WITH_SHORT_NAME = True       # 0 results -> one more search with the first 2 words of the name
CAPTCHA_WAIT_MINUTES = 30          # no captcha typed for this long -> school marked SKIPPED
PAGE_TIMEOUT_MS = 45000            # max wait for a report card to load
BEEP_WHEN_CAPTCHA_NEEDED = True
AUTO_SEARCH_AFTER_CAPTCHA = True   # click Search by itself once you have typed the full captcha
CAPTCHA_LENGTH = 6                 # the site's captcha is 6 characters
BROWSER_ORDER = ("chrome", "msedge", None)   # Google Chrome first; Edge if Chrome missing; None = built-in browser
# ---------------------------------------------------------------------------

ROOT = Path(__file__).resolve().parent.parent
OUTPUT_DIR = ROOT / "output"
PROFILE_DIR = ROOT / "chrome-profile"   # the tool's own Chrome profile (keeps sign-in, cookies)
SITE = "https://kys.udiseplus.gov.in/"
SEARCH_URL = SITE + "#/advancesearch"
ENROL_HEADING = "Enrolment by grade in the current academic session"

ROMAN = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"]
GRADE_MAP = {"PRE-PR": "Pre-Primary", "PRE-PRI": "Pre-Primary", "PRE-PRI.": "Pre-Primary"}
GRADE_MAP.update({r: f"Class {i}" for i, r in enumerate(ROMAN, 1)})
GRADE_COLS = ["Pre-Primary"] + [f"Class {i}" for i in range(1, 13)]
DATA_COLS = ["Academic Year", *GRADE_COLS, "Total Boys", "Total Girls", "Grand Total"]

# columns added to the input sheet: one per grade, then the total and bookkeeping
OUT_COLS = [*GRADE_COLS, "Total Students", "Academic Year", "Status", "Remarks"]
MATCH_COLS = ["Input Row #", "Input School Name", "School Name", "UDISE Code", "Edu. District",
              "Edu. Block", "PIN Code", *DATA_COLS, "Report Card URL", "Remarks"]
NO_DATA_TEXT = "Data not found"     # written in every grade cell when the report card has no data
DONE_STATUSES = {"OK", "MULTIPLE MATCHES", "NOT FOUND", "NO DATA"}
STATUS_FILL = {"OK - CHECK MATCH": "FFEB9C", "OK": "C6EFCE", "NO DATA": "FCE4D6", "MULTIPLE MATCHES": "FFEB9C", "NOT FOUND": "F4CCCC",
               "ERROR": "F4CCCC", "SKIPPED": "D9D9D9"}

COLUMN_ALIASES = {
    "name": ["schoolname", "school", "nameofschool", "name"],
    "state": ["state", "statename"],
    "district": ["district", "edudistrict", "districtname"],
    "pin": ["pincode", "pin", "postalcode"],
    "udise": ["udisecode", "udise", "udiseno", "schoolcode", "udiseschoolcode"],
    "keyword": ["searchkeyword", "keyword", "searchname"],
}


def key(s):
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def norm_name(s):
    return re.sub(r"[^A-Z0-9]", "", str(s or "").upper())


def cell_str(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def log(msg=""):
    print(msg, flush=True)


# ============================ EXCEL ========================================
def find_input_file():
    """The school list = the one .xlsx file in the main folder (next to start.bat)."""
    files = [p for p in ROOT.glob("*.xlsx") if not p.name.startswith("~$")]
    if len(files) == 1:
        return files[0]
    if not files:
        log("ERROR: put your school list (.xlsx) in: " + str(ROOT))
    else:
        log("ERROR: keep only ONE .xlsx file next to start.bat. Found: " + ", ".join(p.name for p in files))
    return None


def read_input(path):
    """Returns (sheet name, header, records)."""
    wb = load_workbook(path, read_only=True, data_only=True)
    sheet = INPUT_SHEET if INPUT_SHEET in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        raise ValueError("The input sheet is empty.")
    header = [cell_str(c) for c in rows[0]]
    while header and not header[-1]:
        header.pop()
    idx = {}
    for field, aliases in COLUMN_ALIASES.items():
        for i, h in enumerate(header):
            if key(h) in aliases:
                idx[field] = i
                break
    if "name" not in idx:
        raise ValueError("Column 'School Name' not found in the first row of " + path.name)

    records = []
    for n, r in enumerate(rows[1:], start=2):
        vals = [cell_str(r[i]) if i < len(r) else "" for i in range(len(header))]
        get = lambda f: vals[idx[f]] if f in idx else ""
        name = get("name")
        if not name:
            continue
        pin = re.sub(r"\D", "", get("pin"))
        records.append({
            "row": n, "values": vals, "name": name,
            "state": get("state") or DEFAULT_STATE,
            "district": get("district"), "pin": pin,
            "udise": re.sub(r"\D", "", get("udise")),
            "keyword": get("keyword") or name,
        })
    return sheet, header, records


class WorkbookStore:
    """Writes results into a COPY of the input workbook (output/<name> - Enrolment.xlsx):
    every sheet and the original columns stay as they are; one column per grade is added
    to the processed sheet, followed by Total Students, Academic Year, Status, Remarks.
    Re-running continues from that copy."""

    def __init__(self, in_path, sheet_name, header, records):
        self.records = records
        self.base = len(header)                       # number of original columns
        self.name_col = 1 + next(i for i, h in enumerate(header) if key(h) in COLUMN_ALIASES["name"])
        self.out_path = OUTPUT_DIR / f"{in_path.stem} - Enrolment.xlsx"
        resuming = self.out_path.exists()
        src = self.out_path if resuming else in_path
        try:
            self.wb = load_workbook(src)
        except Exception as e:     # never overwrite earlier results we could not read
            raise RuntimeError(f"Could not open {src.name} ({e}). Close it in Excel and try again.")
        if sheet_name not in self.wb.sheetnames:
            raise RuntimeError(f"Sheet '{sheet_name}' not found in {src.name}")
        self.ws = self.wb[sheet_name]
        if resuming:   # rows must still line up with the input list
            for rec in records:
                if cell_str(self.ws.cell(rec["row"], self.name_col).value) != rec["name"]:
                    raise RuntimeError(
                        f"The input list changed since '{self.out_path.name}' was started (row {rec['row']}). "
                        f"Move that file out of the 'output' folder to start again.")
        self._setup_columns()
        self.matches = self._read_matches()
        self.unsaved = 0

    def _setup_columns(self):
        ws = self.ws
        existing = {cell_str(c.value): c.column for c in ws[1][self.base:] if c.value is not None}
        self.cols, nxt = {}, self.base + 1
        head_src = ws.cell(1, 1)
        for name in OUT_COLS:
            col = existing.get(name)
            if not col:
                col = max([nxt, *[c + 1 for c in self.cols.values()]])
                cell = ws.cell(1, col, name)
                cell.font = copy(head_src.font)
                cell.fill = copy(head_src.fill)
                cell.border = copy(head_src.border)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                ws.column_dimensions[get_column_letter(col)].width = (
                    45 if name == "Remarks" else 13 if name in ("Status", "Academic Year", "Total Students") else 10)
            self.cols[name] = col
            nxt = col + 1
        last = get_column_letter(max(self.cols.values()))
        if ws.auto_filter.ref:
            ws.auto_filter.ref = f"A1:{last}{ws.max_row}"

    def _read_matches(self):
        if "Matches" not in self.wb.sheetnames:
            return []
        rows = list(self.wb["Matches"].iter_rows(values_only=True))
        if not rows:
            return []
        h = [cell_str(c) for c in rows[0]]
        return [{c: r[h.index(c)] for c in MATCH_COLS if c in h and h.index(c) < len(r)}
                for r in rows[1:] if r and r[0] is not None]

    def is_done(self, rec):
        st = cell_str(self.ws.cell(rec["row"], self.cols["Status"]).value).upper()
        return any(st.startswith(s) for s in DONE_STATUSES)

    def set_result(self, rec, result, match_rows):
        vals = {c: result.get(c) for c in OUT_COLS}
        vals["Total Students"] = result.get("Grand Total")
        for name, col in self.cols.items():
            self.ws.cell(rec["row"], col).value = vals.get(name)
        st_cell = self.ws.cell(rec["row"], self.cols["Status"])
        st = cell_str(vals.get("Status")).upper()
        color = next((c for k, c in STATUS_FILL.items() if st.startswith(k)), None)
        st_cell.fill = PatternFill("solid", fgColor=color) if color else PatternFill(fill_type=None)
        self.matches = [m for m in self.matches if str(m.get("Input Row #")) != str(rec["row"])]
        self.matches.extend(match_rows)
        self.unsaved += 1

    def save(self, force=False):
        if not force and self.unsaved < SAVE_EVERY_N_SCHOOLS:
            return
        if self.matches or "Matches" in self.wb.sheetnames:
            if "Matches" in self.wb.sheetnames:
                del self.wb["Matches"]
            wm = self.wb.create_sheet("Matches")
            wm.append(MATCH_COLS)
            for m in sorted(self.matches, key=lambda d: int(d.get("Input Row #") or 0)):
                wm.append([m.get(c) for c in MATCH_COLS])
            for c in wm[1]:
                c.font = Font(bold=True)
        OUTPUT_DIR.mkdir(exist_ok=True)
        tmp = self.out_path.with_name("~saving.tmp.xlsx")
        while True:
            try:
                self.wb.save(tmp)
                tmp.replace(self.out_path)
                self.unsaved = 0
                return
            except PermissionError:
                log(f"  !! Cannot save - please CLOSE '{self.out_path.name}' in Excel. Retrying in 5 s...")
                time.sleep(5)


# ============================ BROWSER ======================================
JS_BANNER = """([msg, bad]) => {
  let d = document.getElementById('__udise_helper');
  if (!d) {
    d = document.createElement('div'); d.id = '__udise_helper';
    d.style.cssText = 'position:fixed;left:0;right:0;bottom:0;z-index:2147483647;color:#fff;' +
      'font:600 17px/1.4 Segoe UI,Arial,sans-serif;padding:12px 16px;text-align:center;' +
      'box-shadow:0 -2px 10px rgba(0,0,0,.35)';
    document.body.appendChild(d);
  }
  d.style.background = bad ? '#b71c1c' : '#1b5e20';
  d.textContent = msg;
}"""

JS_FIND_TABLE = """
const __findEnrolTable = () => {
  const hs = [...document.querySelectorAll('h1,h2,h3,h4,h5,h6')]
    .filter(h => /Enrolment by grade in the current academic session/i.test(h.textContent));
  for (const h of hs) {
    let a = h;
    for (let i = 0; i < 6 && a; i++) {
      a = a.parentElement;
      if (a) { const t = a.querySelector('table'); if (t) return t; }
    }
  }
  return null;
};
"""

JS_ENROL_READY = "(udise) => {" + JS_FIND_TABLE + """
  if (udise && !document.body.innerText.includes(udise)) return false;
  const t = __findEnrolTable();
  return !!t && [...t.rows].some(r => r.cells.length && /^g\\.?\\s*total$/i.test(r.cells[0].textContent.trim()));
}"""

JS_ENROL_READ = "() => {" + JS_FIND_TABLE + """
  const t = __findEnrolTable();
  if (!t) return null;
  const txt = document.body.innerText;
  const m = txt.match(/Academic Year\\s*\\((\\d{4}-\\d{2})\\)/);
  const nm = txt.match(/School Name[ \\t]+([^\\t\\n]+?)(?=[ \\t]*(?:\\n|\\t|State\\b))/);
  return { year: m ? m[1] : '', name: nm ? nm[1].trim() : '', rows: [...t.rows].map(r => [...r.cells].map(c => c.textContent.trim())) };
}"""

# Reads result cards from the page structure (each card sits in a collapsible accordion panel,
# which Chrome reports as "not visible" even when it is drawn - so nothing here relies on visibility).
JS_CARDS = """() => {
  const clean = s => (s || '').replace(/\\s+/g, ' ').trim();
  const links = [...document.querySelectorAll('a,button')].filter(a => clean(a.textContent) === 'Report Card');
  links.forEach((a, i) => a.setAttribute('data-udise-bot', String(i)));
  return links.map((a, i) => {
    let c = a;
    while (c.parentElement && !/UDISE Code/.test(c.textContent || '')) c = c.parentElement;
    const field = label => {            // <span>Label</span> : <span>value</span>
      const lab = [...c.querySelectorAll('span,b,strong,label')]
        .find(e => clean(e.textContent).replace(/:$/, '').trim().toLowerCase() === label.toLowerCase());
      if (!lab) return '';
      let v = lab.nextElementSibling;
      return v ? clean(v.textContent) : '';
    };
    const t = clean(c.textContent);
    const g = re => { const m = t.match(re); return m ? m[1].trim() : ''; };
    const udEl = c.querySelector('.udiseCode');
    const nameEl = c.querySelector('h1,h2,h3,h4,h5,h6');
    return {
      idx: i, href: a.getAttribute('href') || '',
      udise: udEl ? clean(udEl.textContent) : g(/UDISE Code\\s*:?\\s*(\\d{11})/),
      district: field('Edu. District'),
      block: field('Edu. Block'),
      pin: field('PIN Code') || g(/PIN Code\\s*:?\\s*(\\d{6})/),
      name: nameEl ? clean(nameEl.textContent) : '',
    };
  });
}"""

JS_HAS_CARDS = """() => [...document.querySelectorAll('a,button')].some(a => a.textContent.trim() === 'Report Card')"""


# marks the visible Report Card button that belongs to the card with this UDISE code
JS_MARK_LINK = """(u) => {
  document.querySelectorAll('[data-udise-click]').forEach(x => x.removeAttribute('data-udise-click'));
  const ls = [...document.querySelectorAll('a,button')].filter(a => a.textContent.trim() === 'Report Card');
  const vis = ls.filter(a => a.getClientRects().length > 0 && getComputedStyle(a).visibility !== 'hidden');
  for (const a of [...vis, ...ls.filter(x => !vis.includes(x))]) {   // visible ones first
    let c = a;
    while (c.parentElement && !/UDISE Code/.test(c.textContent || '')) c = c.parentElement;
    if ((c.textContent || '').includes(u)) { a.setAttribute('data-udise-click', '1'); return vis.includes(a); }
  }
  return null;
}"""


def to_int(s):
    s = re.sub(r"[^\d-]", "", str(s))
    return int(s) if s not in ("", "-") else 0


def parse_enrolment(data):
    """Returns (dict of DATA_COLS, remark)."""
    rows = data["rows"]
    header = next((r for r in rows if r and r[0].strip().lower() == "class"), None)
    total = next((r for r in rows if r and r[0].strip().lower() == "total"), None)
    gtotal = next((r for r in rows if r and re.fullmatch(r"g\.?\s*total", r[0].strip().lower())), None)
    if not header or not total:
        raise ValueError("Enrolment table layout not recognised")
    labels = header[1:-1]                      # Pre-Pr, I ... XII  (drop 'Class' and 'Total')
    nums = [to_int(x) for x in total[1:]]
    n = len(labels)
    if len(nums) < 2 * n + 3:
        raise ValueError(f"Enrolment TOTAL row has {len(nums)} numbers, expected {2 * n + 3}")
    out = {"Academic Year": data.get("year", "")}
    remarks = []
    for i, lab in enumerate(labels):
        col = GRADE_MAP.get(lab.strip().upper())
        if not col:
            remarks.append(f"unknown grade column '{lab}'")
            continue
        out[col] = nums[2 * i] + nums[2 * i + 1]
    out["Total Boys"], out["Total Girls"], out["Grand Total"] = nums[2 * n], nums[2 * n + 1], nums[2 * n + 2]
    if gtotal:   # cross-check with the site's own G.Total row
        g = [to_int(x) for x in gtotal[1:]]
        site = [g[i] for i in range(min(n, len(g)))]
        mine = [out.get(GRADE_MAP.get(l.strip().upper(), ""), None) for l in labels]
        if site != mine[:len(site)]:
            remarks.append("grade totals differ from site G.Total row - please verify")
    if sum(out.get(c, 0) for c in GRADE_COLS) != out["Grand Total"]:
        remarks.append("sum of grades != Grand Total - please verify")
    return out, "; ".join(remarks)


class KYS:
    def __init__(self, page):
        self.page = page
        self.search_count = 0
        self.last_search = None
        self.year_seg = ""      # year part of the report-card address, learnt from the first click
        self.nodata_snapshots = 0
        self.social_count = 0   # report-card data calls seen (used to wait for real numbers)
        self.rc_count = 0
        self.region_count = 0   # Advance Search (district) responses seen
        self.last_region = None
        page.on("response", self._on_response)

    def _on_response(self, r):
        u = r.url
        if "search-schools" in u:
            self.search_count += 1
            self.last_search = r
        elif "search-school/by-region" in u:
            self.region_count += 1
            self.last_region = r
        elif "getSocialData" in u:
            self.social_count += 1
        elif "school/report-card" in u:
            self.rc_count += 1

    def save_debug(self, tag):
        """Saves the page (HTML + screenshot) to output/debug to diagnose problems."""
        try:
            d = OUTPUT_DIR / "debug"
            d.mkdir(parents=True, exist_ok=True)
            stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            (d / f"{stamp}_{tag}.html").write_text(self.page.content(), encoding="utf-8")
            self.page.screenshot(path=str(d / f"{stamp}_{tag}.png"), full_page=True)
            log(f"  (debug snapshot saved: output/debug/{stamp}_{tag}.*)")
        except Exception:
            pass

    def banner(self, msg, bad=False):
        try:
            self.page.evaluate(JS_BANNER, [msg, bad])
        except PWError:
            pass

    # ---- search page ----
    def open_search(self):
        """Full page load of the search page. The State list only loads on a real page load,
        and the site empties it after a wrong captcha - so always reload."""
        p = self.page
        for attempt in range(4):
            try:
                if attempt % 2 == 0:
                    p.goto(SEARCH_URL, wait_until="domcontentloaded")
                    p.reload(wait_until="domcontentloaded")
                else:  # fallback: home page first, then the search page
                    p.goto(SITE, wait_until="domcontentloaded")
                    p.reload(wait_until="domcontentloaded")
                    p.wait_for_timeout(2500)
                    p.goto(SEARCH_URL, wait_until="domcontentloaded")
                p.locator("#school").wait_for(state="visible", timeout=20000)
                p.locator("#school").check()
                p.wait_for_function(
                    "() => [...document.querySelectorAll('select')].some(s => s.options.length > 5)",
                    timeout=15000)
                return
            except PWError as e:
                if "closed" in str(e).lower():
                    raise
                log("  (search page not ready, retrying...)")
        raise RuntimeError("Search page did not load (State list empty). Check internet.")

    def fill_form(self, state, keyword):
        p = self.page
        if not p.locator("#school").is_checked():
            p.locator("#school").check()
        sel = p.locator("select").filter(has=p.locator("option", has_text=re.compile(r"MAHARASHTRA", re.I))).first
        labels = [s.strip() for s in sel.locator("option").all_inner_texts()]
        wanted = next((l for l in labels if l.upper() == state.strip().upper()), None)
        if not wanted:
            wanted = next((l for l in labels if state.strip().upper() in l.upper()), None)
        if not wanted:
            raise ValueError(f"State '{state}' is not in the site's State list")
        sel.select_option(label=wanted)
        name_box = p.locator('input[type="text"]:visible:not([placeholder="Enter Captcha"]):not([placeholder="Search"])').first
        name_box.fill("")
        name_box.fill(keyword)
        name_box.press("Tab")
        cap = p.locator('input[placeholder="Enter Captcha"]').first
        cap.fill("")
        cap.scroll_into_view_if_needed()
        cap.focus()

    def invalid_captcha_shown(self):
        try:
            return self.page.get_by_text("Invalid Captcha").first.is_visible()
        except PWError:
            return False

    def search_with_captcha(self, state, keyword, progress):
        """Fills the form and waits for the user to solve the captcha. Returns True when results are shown."""
        p = self.page
        self.open_search()
        self.fill_form(state, keyword)
        p.bring_to_front()
        how = "search starts automatically" if AUTO_SEARCH_AFTER_CAPTCHA else "press ENTER"
        msg = f"{progress}  |  Searching: {keyword}  ->  Type the CAPTCHA above ({how})"
        self.banner(msg)
        log(f"  Waiting for captcha...  (search: '{keyword}', state: {state})")
        if BEEP_WHEN_CAPTCHA_NEEDED:
            try:
                import winsound
                winsound.MessageBeep()
            except Exception:
                pass
        start = self.search_count
        submitted = None
        cap = p.locator('input[placeholder="Enter Captcha"]').first
        search_btn = p.locator("button", has_text=re.compile(r"^\s*Search\s*$")).first
        deadline = time.time() + CAPTCHA_WAIT_MINUTES * 60
        while time.time() < deadline:
            p.wait_for_timeout(300)
            if self.search_count == start:
                if AUTO_SEARCH_AFTER_CAPTCHA:
                    try:
                        val = cap.input_value().strip()
                    except PWError:
                        val = ""
                    if len(val) >= CAPTCHA_LENGTH and val != submitted:
                        submitted = val          # you typed it - now press Search for you
                        try:
                            search_btn.click()
                        except PWError:
                            pass
                continue
            p.wait_for_timeout(1500)   # let the page render the result
            body = p.inner_text("body")
            shown = re.search(r"Showing Result For School Name", body, re.I)
            if not shown and self.invalid_captcha_shown():
                log("  Wrong captcha - please try again.")
                self.open_search()               # the site clears the form AND the State list
                self.fill_form(state, keyword)
                self.banner(f"{progress}  |  WRONG CAPTCHA - type the new captcha", bad=True)
                start = self.search_count
                submitted = None
                continue
            return True
        return False

    def read_cards(self):
        p = self.page
        body = p.inner_text("body")
        m = re.search(r"Showing\s+(\d+)\s+Result", body)
        total = int(m.group(1)) if m else 0
        if total > 10:   # show up to 100 results on one page
            try:
                size = p.locator("select").filter(has=p.locator("option", has_text=re.compile(r"^\s*100\s*$"))).first
                size.select_option(label="100")
                p.wait_for_timeout(1500)
            except PWError:
                pass
        if total:
            try:   # result cards appear a moment after the search finishes
                p.wait_for_function(JS_HAS_CARDS, timeout=15000)
                p.wait_for_timeout(500)
            except PWError:
                self.save_debug("no_cards")
        cards = p.evaluate(JS_CARDS)
        ids = self._school_ids()
        for c in cards:
            c["school_id"] = ids.get(c["udise"], "")
        # schools beyond the results shown on the page are still in the search data the page got
        shown = {c["udise"] for c in cards}
        for ud, sid in ids.items():
            if ud not in shown:
                cards.append({"idx": -1, "href": "", "udise": ud, "school_id": sid,
                              "name": "", "district": "", "block": "", "pin": ""})
        return cards, max(total, len(cards))

    # ---- bulk: one Advance Search for a whole district ----
    def open_advance(self, state, district):
        """Search page (full load) -> Advance Search -> State + District selected, cursor in captcha."""
        p = self.page
        for attempt in range(3):
            try:
                if attempt % 2 == 0:
                    p.goto(SEARCH_URL, wait_until="domcontentloaded")
                    p.reload(wait_until="domcontentloaded")
                else:   # fallback: home page first (loads the State list), then the search page
                    p.goto(SITE, wait_until="domcontentloaded")
                    p.reload(wait_until="domcontentloaded")
                    p.wait_for_timeout(2000)
                    p.goto(SEARCH_URL, wait_until="domcontentloaded")
                p.get_by_text("Advance Search").first.click(timeout=20000)
                # the State list appears only once the Advance Search panel is open
                p.wait_for_function("() => [...document.querySelectorAll('select')].some(s => "
                                    "s.getClientRects().length && [...s.options].some(o => o.text.trim() === 'MAHARASHTRA'))",
                                    timeout=20000)
                st = p.locator("select:visible").filter(
                    has=p.locator("option", has_text=re.compile(rf"^\s*{re.escape(state)}\s*$", re.I))).first
                st.select_option(label=next(l.strip() for l in st.locator("option").all_inner_texts()
                                            if l.strip().upper() == state.upper()))
                p.wait_for_function("(d) => [...document.querySelectorAll('select')].some(s => "
                                    "[...s.options].some(o => o.text.trim().toUpperCase() === d))",
                                    arg=district.upper(), timeout=20000)
                dt = p.locator("select:visible").filter(
                    has=p.locator("option", has_text=re.compile(rf"^\s*{re.escape(district)}\s*$", re.I))).first
                dt.select_option(label=next(l.strip() for l in dt.locator("option").all_inner_texts()
                                            if l.strip().upper() == district.upper()))
                cap = p.locator('input[placeholder="Enter Captcha"]:visible').first
                cap.fill("")
                cap.scroll_into_view_if_needed()
                cap.focus()
                return
            except PWError as e:
                if "closed" in str(e).lower():
                    raise
                log("  (advance search not ready, retrying...)")
        raise RuntimeError("Advance Search form did not load. Check internet.")

    def region_search_with_captcha(self, state, district):
        """One captcha -> list of every school in the district. Returns {udise: school_id}."""
        p = self.page
        self.open_advance(state, district)
        p.bring_to_front()
        msg = f"ONE-TIME district search: {state} / {district}  ->  type the CAPTCHA above (search starts automatically)"
        self.banner(msg)
        log(f"  Waiting for captcha...  (Advance Search: {state} / {district})")
        if BEEP_WHEN_CAPTCHA_NEEDED:
            try:
                import winsound
                winsound.MessageBeep()
            except Exception:
                pass
        start, submitted = self.region_count, None
        deadline = time.time() + CAPTCHA_WAIT_MINUTES * 60
        while time.time() < deadline:
            p.wait_for_timeout(300)
            if self.region_count == start:
                if AUTO_SEARCH_AFTER_CAPTCHA:
                    try:
                        val = p.locator('input[placeholder="Enter Captcha"]:visible').first.input_value().strip()
                    except PWError:
                        val = ""
                    if len(val) >= CAPTCHA_LENGTH and val != submitted:
                        submitted = val
                        try:
                            p.locator("button:visible", has_text=re.compile(r"^\s*Search\s*$")).first.click()
                        except PWError:
                            pass
                continue
            p.wait_for_timeout(2500)          # big result - let the page take it in
            ids = self._school_ids(self.last_region)
            if not ids and self.invalid_captcha_shown():
                log("  Wrong captcha - please try again.")
                self.open_advance(state, district)
                self.banner("WRONG CAPTCHA - type the new captcha (district search)", bad=True)
                start, submitted = self.region_count, None
                continue
            if not self.year_seg:
                try:
                    self._learn_year_seg()     # one Report Card click shows the report-card address format
                except PWError:
                    pass
            return ids
        return None

    def _school_ids(self, resp=None):
        """UDISE code -> site's internal school id, taken from the search result the page received."""
        out = {}
        resp = resp or self.last_search
        try:
            js = resp.json() if resp else None
        except Exception:
            return out

        def walk(o):
            if isinstance(o, dict):
                sid = next((o[k] for k in o if re.fullmatch(r"school_?id", str(k), re.I)), None)
                if sid not in (None, ""):
                    for v in o.values():
                        if isinstance(v, (str, int)) and re.fullmatch(r"\d{11}", str(v).strip()):
                            out[str(v).strip()] = str(sid)
                for v in o.values():
                    walk(v)
            elif isinstance(o, list):
                for v in o:
                    walk(v)
        walk(js)
        return out

    # ---- report card ----
    def _wait_report_data(self, s0, r0):
        """The report card first shows a table full of zeros; the real numbers arrive with the
        page's own data calls (school/report-card + getSocialData). Wait until those have
        arrived and stopped, so we never read the placeholder zeros."""
        p = self.page
        deadline = time.time() + PAGE_TIMEOUT_MS / 1000
        last_n, last_change = -1, time.time()
        while time.time() < deadline:
            p.wait_for_timeout(250)
            n = self.social_count - s0
            if n != last_n:
                last_n, last_change = n, time.time()
            if self.rc_count > r0 and n >= 1 and time.time() - last_change >= 1.5:
                return True
        return False

    def _stable_read(self):
        """Reads the enrolment table until two reads 0.8 s apart are identical."""
        p = self.page
        prev = p.evaluate(JS_ENROL_READ)
        for _ in range(8):
            p.wait_for_timeout(800)
            cur = p.evaluate(JS_ENROL_READ)
            if cur and prev and cur["rows"] == prev["rows"]:
                return cur
            prev = cur
        return prev

    def _learn_year_seg(self):
        """Clicks any Report Card button once to learn the year part of report-card addresses."""
        p = self.page
        if p.evaluate("() => { const a = [...document.querySelectorAll('a,button')]"
                      ".find(x => x.textContent.trim() === 'Report Card'); if (a) a.click(); return !!a; }"):
            p.wait_for_url(re.compile(r"reportcard"), timeout=PAGE_TIMEOUT_MS)
            m = re.search(r"#/reportcard/\d+/(\d+)", p.url)
            if m:
                self.year_seg = m.group(1)

    def fetch_report(self, cand):
        p = self.page
        s0, r0 = self.social_count, self.rc_count
        on_page = cand.get("idx", -1) >= 0
        if not cand.get("href") and not on_page and cand.get("school_id") and not self.year_seg:
            self._learn_year_seg()
        direct = ""
        if cand.get("href"):
            direct = urljoin(SITE, cand["href"])
        elif on_page and p.evaluate(JS_MARK_LINK, cand.get("udise") or "-") is not None:
            direct = ""          # its card is on the results page -> click its Report Card button
        elif cand.get("school_id") and self.year_seg:
            # same address the site's own Report Card button opens
            direct = f"{SITE}#/reportcard/{cand['school_id']}/{self.year_seg}"
        if direct:
            # full reload: going from one report card straight to another does not refresh the page
            p.goto(direct, wait_until="domcontentloaded")
            p.reload(wait_until="domcontentloaded")
        else:
            found = p.evaluate(JS_MARK_LINK, cand["udise"]) if cand.get("udise") else None
            if found is None:
                self.save_debug(f"no_link_{cand.get('udise')}")
                raise ValueError("Report Card button for this school not found on the results page")
            # the card sits in a collapsible panel Chrome calls "not visible", so click from inside the page
            p.locator('[data-udise-click="1"]').first.evaluate("a => a.click()")
            p.wait_for_url(re.compile(r"reportcard"), timeout=PAGE_TIMEOUT_MS)
            m = re.search(r"#/reportcard/\d+/(\d+)", p.url)
            if m:
                self.year_seg = m.group(1)
        udise = cand.get("udise") or ""
        # 1) wait until THIS school's report card is on screen (its UDISE code is shown),
        #    or the site's "Oops! Data Not Found : 2025-26" popup appears
        for attempt in range(2):        # the site is sometimes slow: reload once, then give up
            try:
                p.wait_for_function("(u) => { const t = document.body.innerText; return /Data Not Found/i.test(t) || "
                                    "(t.includes(u) && /School Report Card/i.test(t)); }",
                                    arg=udise, timeout=25000 if attempt == 0 else PAGE_TIMEOUT_MS)
                nf = re.search(r"Data Not Found\s*:?\s*(\d{4}-\d{2})?", p.inner_text("body"), re.I)
                if nf:
                    try:   # close the popup (OK)
                        p.get_by_role("button", name=re.compile(r"^\s*OK\s*$", re.I)).first.click(timeout=2000)
                    except PWError:
                        pass
                    out = {c: NO_DATA_TEXT for c in GRADE_COLS}
                    out.update({"Grand Total": NO_DATA_TEXT, "Academic Year": nf.group(1) or "", "_nodata": True})
                    return out, f"Site says: Data Not Found{(' : ' + nf.group(1)) if nf.group(1) else ''}", p.url
                break
            except PWError as e:
                if "closed" in str(e).lower():
                    raise
                if attempt:
                    self.save_debug(f"report_{udise}")
                    raise ValueError(f"Report card did not load: {p.url}")
                log("  (report card slow - reloading once)")
                s0, r0 = self.social_count, self.rc_count
                p.reload(wait_until="domcontentloaded")
        # 2) wait for the page's own data calls to finish
        loaded = self._wait_report_data(s0, r0)
        # 3) enrolment table present? (a few extra seconds for it to draw)
        try:
            p.wait_for_function(JS_ENROL_READY, arg=udise, timeout=8000 if loaded else 20000)
        except PWError as e:
            if "closed" in str(e).lower():
                raise
            if not loaded:
                raise ValueError("Report card data did not finish loading (slow site) - will retry next run")
            # page fully loaded but has no grade-wise enrolment table -> the site has no data
            if self.nodata_snapshots < 5:
                self.nodata_snapshots += 1
                self.save_debug(f"nodata_{udise}")
            yr = re.search(r"Academic Year\s*\((\d{4}-\d{2})\)", p.inner_text("body"))
            out = {c: NO_DATA_TEXT for c in GRADE_COLS}
            out.update({"Grand Total": NO_DATA_TEXT, "Academic Year": yr.group(1) if yr else "", "_nodata": True})
            return out, "Report card has no grade-wise enrolment data", p.url
        data = self._stable_read()
        if not data:
            raise ValueError("Enrolment section not found on report card")
        if data.get("name") and not cand.get("name"):
            cand["name"] = data["name"]           # name as printed on the report card
        out, remark = parse_enrolment(data)
        return out, remark, p.url


def choose(cands, rec):
    c = list(cands)
    if rec["udise"]:
        c = [x for x in c if x["udise"] == rec["udise"]]
    if rec["district"]:
        d = norm_name(rec["district"])
        c = [x for x in c if d in norm_name(x["district"])]
    if rec["pin"]:
        c = [x for x in c if x["pin"] == rec["pin"]]
    if len(c) > 1:
        exact = [x for x in c if norm_name(x["name"]) == norm_name(rec["name"])]
        if len(exact) == 1:
            c = exact
    return c


def short_keyword(name):
    words = [w for w in re.sub(r"[^A-Za-z0-9 ]", " ", name).split() if w.lower() != "the"]
    kw = " ".join(words[:2])
    return kw if len(kw) >= 3 else ""


def process_school(kys, rec, progress):
    keywords = [rec["keyword"]]
    if RETRY_WITH_SHORT_NAME:
        sk = short_keyword(rec["keyword"])
        if sk and sk.lower() != rec["keyword"].lower():
            keywords.append(sk)

    cands, chosen, total, used_kw = [], [], 0, rec["keyword"]
    for kw in keywords:
        if not kys.search_with_captcha(rec["state"], kw, progress):
            return {"Status": "SKIPPED", "Remarks": f"No captcha entered within {CAPTCHA_WAIT_MINUTES} min"}, []
        cands, total = kys.read_cards()
        used_kw = kw
        chosen = choose(cands, rec)
        log(f"  Site returned {total} result(s) for '{kw}'"
            + (f" - UDISE {rec['udise']} {'found' if chosen else 'not among them'}" if rec["udise"] else ""))
        if chosen:      # with a UDISE code: retry the shorter name if the code was not in the results
            break

    if not cands:
        return {"Status": "NOT FOUND", "Matches Found": 0,
                "Remarks": f"No school found for '{' / '.join(keywords)}' in {rec['state']}"}, []

    kw_note = "" if used_kw == rec["keyword"] else f"found by searching '{used_kw}'"
    if not chosen:
        named = [x for x in cands if x["name"]]
        others = "; ".join(f"{x['name']} ({x['udise']}, {x['district']})" for x in named[:5])
        what = "UDISE code" if rec["udise"] else "District/PIN"
        return {"Status": "NOT FOUND", "Matches Found": 0,
                "Remarks": f"Searched '{' / '.join(keywords)}': {len(cands)} result(s), none with this {what}. "
                           f"Site shows e.g.: {others}"}, []

    fetched = []
    for cand in chosen[:MAX_CANDIDATES_TO_FETCH]:
        try:
            data, remark, url = kys.fetch_report(cand)
            fetched.append((cand, data, remark, url))
            log(f"  -> {cand['name']} [{cand['udise']}] total={data.get('Grand Total')}")
        except Exception as e:
            first_line = (str(e).strip().splitlines() or ["unknown error"])[0][:200]
            fetched.append((cand, {}, f"ERROR reading report card: {first_line}", kys.page.url))
            log(f"  -> {cand['name']} [{cand['udise']}] ERROR: {first_line}")

    def base(cand):
        return {"Matched School Name": cand["name"], "Matched UDISE Code": cand["udise"],
                "Matched Edu. District": cand["district"], "Matched Edu. Block": cand["block"],
                "Matched PIN Code": cand["pin"]}

    if len(chosen) == 1:
        cand, data, remark, url = fetched[0]
        res = {**base(cand), **data, "Matches Found": 1, "Report Card URL": url,
               "Remarks": "; ".join(x for x in (remark, kw_note) if x)}
        res["Status"] = "ERROR" if remark.startswith("ERROR") else "NO DATA" if data.pop("_nodata", False) else "OK"
        res.pop("_nodata", None)
        if res["Status"] == "OK" and used_kw != rec["keyword"] and not rec["udise"]:
            res["Status"] = "OK - CHECK MATCH"
            res["Remarks"] = ("Full name not found; matched using shorter name - confirm this is the right "
                              "school (or add its UDISE Code). " + res["Remarks"])
        return res, []

    match_rows = []
    for cand, data, remark, url in fetched:
        match_rows.append({"Input Row #": rec["row"], "Input School Name": rec["name"],
                           "School Name": cand["name"], "UDISE Code": cand["udise"],
                           "Edu. District": cand["district"], "Edu. Block": cand["block"],
                           "PIN Code": cand["pin"], **data, "Report Card URL": url, "Remarks": remark})
    extra = f"; only first {MAX_CANDIDATES_TO_FETCH} fetched" if len(chosen) > MAX_CANDIDATES_TO_FETCH else ""
    return {"Status": "MULTIPLE MATCHES", "Matches Found": len(chosen),
            "Remarks": f"{len(chosen)} schools matched - see 'Matches' sheet. Add District / PIN / UDISE Code "
                       f"in your list to pick one{extra}. {kw_note}".strip()}, match_rows


class BrowserClosed(Exception):
    def __init__(self, rec):
        super().__init__("browser closed")
        self.rec = rec


def is_closed(e):
    s = str(e).lower()
    return "closed" in s or "target page" in s or "target crashed" in s


ID_CACHE_FILE = OUTPUT_DIR / "school_ids.json"


def load_id_cache():
    """District list saved by an earlier run: ({udise: school_id}, year_seg) or ({}, '')."""
    try:
        d = json.loads(ID_CACHE_FILE.read_text(encoding="utf-8"))
        made = datetime.datetime.fromisoformat(d["created"])
        if (d.get("state"), d.get("district")) != (BULK_STATE, BULK_DISTRICT):
            return {}, ""
        if (datetime.datetime.now() - made).days >= ID_CACHE_DAYS:
            return {}, ""
        return d["ids"], d.get("year_seg", "")
    except Exception:
        return {}, ""


def save_id_cache(ids, year_seg):
    OUTPUT_DIR.mkdir(exist_ok=True)
    ID_CACHE_FILE.write_text(json.dumps({
        "created": datetime.datetime.now().isoformat(timespec="seconds"),
        "state": BULK_STATE, "district": BULK_DISTRICT, "year_seg": year_seg, "ids": ids}), encoding="utf-8")


def bulk_fetch(kys, rec, school_id):
    """Opens one school's report card directly (no search, no captcha) and returns its result row."""
    cand = {"idx": -1, "href": "", "udise": rec["udise"], "school_id": school_id,
            "name": rec["name"], "district": "", "block": "", "pin": ""}
    try:
        data, remark, url = kys.fetch_report(cand)
    except PWError as e:
        if is_closed(e):
            raise BrowserClosed(rec)
        return {"Status": "ERROR", "Remarks": "ERROR reading report card: " + str(e).strip().splitlines()[0][:200]}
    except Exception as e:
        return {"Status": "ERROR", "Remarks": "ERROR reading report card: " + (str(e).strip().splitlines() or ["?"])[0][:200]}
    nodata = data.pop("_nodata", False)
    return {**data, "Status": "NO DATA" if nodata else "OK", "Remarks": remark}


def launch_browser(p):
    """Opens a normal (non-private) Chrome window with its own saved profile in the
    'chrome-profile' folder. (Chrome does not allow automation of your main everyday profile.)"""
    PROFILE_DIR.mkdir(exist_ok=True)
    # chromium_sandbox=True keeps Chrome's security sandbox ON (no "--no-sandbox" warning bar)
    kw = {"headless": False, "args": ["--start-maximized"], "no_viewport": True, "chromium_sandbox": True}
    errors = []
    for channel in BROWSER_ORDER:
        try:
            opts = dict(kw)
            if channel:
                opts["channel"] = channel
            return p.chromium.launch_persistent_context(str(PROFILE_DIR), **opts)
        except Exception as e:
            errors.append(f"{channel}: {str(e).splitlines()[0][:150]}")
            if "user data directory is already in use" in str(e).lower() or "ProcessSingleton" in str(e):
                log("The tool's Chrome window is already open (from another run). Close it and try again.")
                raise
    log("Could not start Chrome/Edge: " + " | ".join(errors))
    log("Downloading a built-in browser (one time)...")
    subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=False)
    return p.chromium.launch_persistent_context(str(PROFILE_DIR), **kw)


def main():
    log("=" * 70)
    log(" UDISE+ Know Your School - Grade-wise Enrolment Collector")
    log("=" * 70)
    path = find_input_file()
    if not path:
        return 1
    try:
        sheet, header, records = read_input(path)
    except Exception as e:
        log(f"ERROR reading {path.name}: {e}")
        return 1
    try:
        store = WorkbookStore(path, sheet, header, records)
    except RuntimeError as e:
        log(f"ERROR: {e}")
        return 1
    pending = [r for r in records if not store.is_done(r)]
    log(f"Input: {path.name} [sheet '{sheet}']  |  schools: {len(records)}  |  already done: "
        f"{len(records) - len(pending)}  |  to do: {len(pending)}")
    log(f"Output: {store.out_path}")
    if not pending:
        store.save(force=True)
        log("Nothing to do - all schools already processed.")
        return 0
    if len(pending) < len(records):
        log(f"\n>> RESUMING from row {pending[0]['row']}: {pending[0]['name']}"
            f"  ({len(records) - len(pending)} schools already saved)")
    if BULK_MODE:
        log("\nA browser window will open. BULK MODE: at most one captcha for the whole district,"
            "\nthen schools are read automatically. Leave the window open.")
    else:
        log("\nA browser window will open. For each school just type the CAPTCHA"
            + (" - Search is clicked automatically." if AUTO_SEARCH_AFTER_CAPTCHA else " and press ENTER."))
    log("Press Ctrl+C here to stop at any time (progress is saved).\n")

    stats = {"OK": 0, "MULTIPLE MATCHES": 0, "NOT FOUND": 0, "ERROR": 0, "SKIPPED": 0}
    with sync_playwright() as pw:
        ctx = launch_browser(pw)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.set_default_timeout(30000)
        kys = KYS(page)
        try:
            leftover = pending
            # ---------- BULK: one captcha for the district, then automatic report cards ----------
            if BULK_MODE and any(r["udise"] for r in pending):
                ids, year = load_id_cache()
                if ids:
                    log(f"Using the saved district list ({len(ids)} schools) - no captcha needed.")
                else:
                    log(f"STEP 1 - ONE captcha lists every school in {BULK_DISTRICT} district.")
                    try:
                        ids = kys.region_search_with_captcha(BULK_STATE, BULK_DISTRICT) or {}
                    except PWError as e:
                        if is_closed(e):
                            raise BrowserClosed(pending[0])
                        log(f"  District search problem: {str(e).splitlines()[0][:150]}")
                        ids = {}
                    except RuntimeError as e:
                        log(f"  District search problem: {e}")
                        ids = {}
                    if ids:
                        save_id_cache(ids, kys.year_seg)
                        log(f"  District list received: {len(ids)} schools.")
                    else:
                        log("  District search did not work - continuing with name search for every school.")
                kys.year_seg = kys.year_seg or year or DEFAULT_YEAR_SEG
                bulk = [r for r in pending if r["udise"] in ids]
                leftover = [r for r in pending if r["udise"] not in ids]
                if bulk:
                    log(f"\nSTEP 2 - {len(bulk)} schools are read AUTOMATICALLY (no captcha). Leave the window open."
                        + (f"\n         {len(leftover)} school(s) not in the district list get a name search after."
                           if leftover else "") + "\n")
                for i, rec in enumerate(bulk, 1):
                    log(f"[auto {i}/{len(bulk)}] Row {rec['row']}: {rec['name']}")
                    result = bulk_fetch(kys, rec, ids[rec["udise"]])   # raises BrowserClosed
                    store.set_result(rec, result, [])
                    if store.unsaved >= BULK_SAVE_EVERY:
                        store.save(force=True)
                    st = result.get("Status", "ERROR")
                    stats[st] = stats.get(st, 0) + 1
                    log(f"  {st}: total={result.get('Grand Total')}" if st in ("OK", "NO DATA")
                        else f"  {st} - {result.get('Remarks')}")
                    kys.banner(f"AUTOMATIC: {i} of {len(bulk)} done - no captcha needed, leave this window open")
                    page.wait_for_timeout(int(BULK_DELAY_SECONDS * 1000))
                store.save(force=True)
                if leftover:
                    log(f"\nSTEP 3 - name search (captcha) for {len(leftover)} school(s) not in the district list.\n")
            # ---------- NAME SEARCH: one captcha per school ----------
            for i, rec in enumerate(leftover, 1):
                progress = f"School {i} of {len(leftover)}"
                log(f"[{i}/{len(leftover)}] Row {rec['row']}: {rec['name']}")
                try:
                    result, match_rows = process_school(kys, rec, progress)
                except PWError as e:
                    if is_closed(e):
                        raise BrowserClosed(rec)
                    result, match_rows = {"Status": "ERROR", "Remarks": str(e)[:300]}, []
                except BrowserClosed:
                    raise
                except Exception as e:
                    result, match_rows = {"Status": "ERROR", "Remarks": str(e)[:300]}, []
                store.set_result(rec, result, match_rows)
                store.save()
                st = result.get("Status", "ERROR")
                stats[st] = stats.get(st, 0) + 1
                log(f"  Status: {st}" + (f" - {result.get('Remarks')}" if st != "OK" else "") + "\n")
        except BrowserClosed as bc:
            log("\nBrowser was closed - stopping. Progress is saved.")
            log(f">> Next time start.bat continues from row {bc.rec['row']}: {bc.rec['name']}")
        except KeyboardInterrupt:
            log("\nStopped by user. Progress is saved - run start.bat again to continue.")
        finally:
            store.save(force=True)
            try:
                ctx.close()
            except Exception:
                pass
    log("-" * 70)
    log("Summary this run: " + ", ".join(f"{k}: {v}" for k, v in stats.items()))
    log(f"Results saved in: {store.out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
