UDISE+ KNOW YOUR SCHOOL - GRADE-WISE ENROLMENT COLLECTOR
=======================================================

WHAT IT DOES
For every school in your list it searches https://kys.udiseplus.gov.in by School Name,
picks the result whose UDISE code matches the list's "School Code" column, opens the
school's Report Card and copies "Enrolment by grade in the current academic session".

BULK MODE (default): at the start you type ONE captcha - the site's Advance Search then lists
every school in Pune district (8,000+). After that every school's Report Card is opened
automatically, with NO more captchas. Only the few schools that are not in the district list
(e.g. Army / Air Force schools) get a normal name search, one captcha each, at the end.
The district list is saved (output\school_ids.json) and reused for 7 days, so stopping and
starting again needs no captcha at all. Search is clicked automatically once 6 characters
of a captcha are typed.


FILES
  start.bat                       double-click to run
  Pune District List.xlsx         your list (keep exactly ONE .xlsx file next to start.bat)
  output\Pune District List - Enrolment.xlsx    the result (created on first run)
  app\                            the program
  sample\                         earlier test files - not used, safe to delete


HOW TO RUN
1. Double-click  start.bat
   - First run on a PC: installs Python (if missing) and the libraries. Needs internet, 1-2 min.
2. A Chrome window opens on the UDISE site. A green bar at the bottom shows which school is
   being searched. Type the captcha. A red bar = wrong captcha: type the new one.
3. STOP ANY TIME: close the Chrome window, close the black window, or press Ctrl+C.
   Every finished school is saved immediately. Next time you double-click start.bat it shows
   ">> RESUMING from row ..." and carries on with the first school that has no result yet
   (the school you were on when you stopped is simply searched again).


THE RESULT (output\Pune District List - Enrolment.xlsx)
A copy of your workbook - both sheets and your original columns unchanged. In the
"Focused Data" sheet these columns are added:
  Pre-Primary, Class 1 ... Class 12   students in each grade
  Total Students                      all grades together
  Academic Year                       year of the data shown on the report card
  Status / Remarks:
    OK          numbers filled in (school confirmed by UDISE code)
    NO DATA     the site shows "Data Not Found" for the school; every grade cell says
                "Data not found"
    NOT FOUND   name search did not return this UDISE code (also tried a shorter name);
                Remarks shows what the site returned. Fix the name in your list and redo.
    ERROR / SKIPPED   retried automatically on the next run
  Your original "Pune District List.xlsx" is never changed.

REDO A SCHOOL: clear its Status cell in the output file, save, close Excel, run start.bat.
Keep the output file CLOSED in Excel while the tool runs (it saves after every school).
Do not insert/delete/sort rows in your list after starting - the tool will refuse to mix them up.
To start over: move the output file out of the output folder.


INSTALL ON ANOTHER PC (from GitHub)
  1. Open https://github.com/ayushhbhuutada/udise-enrolment-automation
  2. Green "Code" button -> "Download ZIP". Unzip it anywhere.
  3. Double-click start.bat.
  To continue from where another PC stopped, copy that PC's "output" folder into this folder.
  Needs: Windows 10/11, internet, Google Chrome (or Microsoft Edge).

UPDATES
  start.bat tells you when a newer version is on GitHub. Double-click update.bat to install it.
  update.bat replaces only the program files - your output (results), chrome-profile and
  your .xlsx list are never touched.


SETTINGS (optional) - top of app\kys_enrolment.py
  INPUT_SHEET (default "Focused Data"), DEFAULT_STATE, RETRY_WITH_SHORT_NAME,
  AUTO_SEARCH_AFTER_CAPTCHA, CAPTCHA_LENGTH, BROWSER_ORDER, CAPTCHA_WAIT_MINUTES


NOTES
  - Data is as published by UDISE+ on the school's Report Card (latest year shown there).
  - Each school's grade totals are cross-checked with the site's own G.Total row; any
    mismatch is written in Remarks.
  - output\debug\ holds page snapshots saved when something unexpected happens. Safe to delete.
  - Google sign-in is blocked by Google inside automated browsers. It is not needed.
