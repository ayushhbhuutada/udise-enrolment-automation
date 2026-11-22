# Updates the tool from GitHub. Never touches your data:
#   output\ (results), chrome-profile\, .venv\ and your .xlsx list are left as they are.
#   -CheckOnly : only tells whether a newer version exists (used by start.bat)
param([switch]$CheckOnly, [switch]$Force)

$ErrorActionPreference = 'Stop'
$Repo   = 'ayushhbhuutada/udise-enrolment-automation'
$Branch = 'main'
$Root   = Split-Path -Parent $PSScriptRoot
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$local = '0'
if (Test-Path "$Root\VERSION") { $local = (Get-Content "$Root\VERSION" -Raw).Trim() }

try {
    $stamp  = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $remote = (Invoke-WebRequest -UseBasicParsing -TimeoutSec 10 `
               "https://raw.githubusercontent.com/$Repo/$Branch/VERSION?t=$stamp").Content.Trim()
} catch {
    if ($CheckOnly) { exit 0 }          # offline: just run the tool
    Write-Host "Could not reach GitHub: $($_.Exception.Message)"
    exit 1
}

if ($CheckOnly) {
    if ($remote -ne $local) {
        Write-Host ""
        Write-Host "  *** A new version ($remote) is available - you have $local. ***"
        Write-Host "  *** Close this window and double-click update.bat to install it. ***"
        Write-Host ""
    }
    exit 0
}

Write-Host "Installed version: $local    Latest version on GitHub: $remote"
if ($remote -eq $local -and -not $Force) {
    Write-Host "You already have the latest version."
    exit 0
}

$tmp = Join-Path $env:TEMP ("udise_update_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp | Out-Null
try {
    Write-Host "Downloading..."
    $zip = Join-Path $tmp 'update.zip'
    Invoke-WebRequest -UseBasicParsing -TimeoutSec 120 "https://github.com/$Repo/archive/refs/heads/$Branch.zip" -OutFile $zip
    Expand-Archive -Path $zip -DestinationPath $tmp
    $src = (Get-ChildItem $tmp -Directory | Select-Object -First 1).FullName

    # program files only
    New-Item -ItemType Directory -Force -Path "$Root\app" | Out-Null
    Copy-Item "$src\app\*" "$Root\app" -Recurse -Force
    foreach ($f in 'start.bat', 'update.bat', 'README.txt', 'VERSION', '.gitignore') {
        if (Test-Path "$src\$f") { Copy-Item "$src\$f" "$Root\$f" -Force }
    }
    # school list: only added when this folder has none (your own copy is never replaced)
    $mine = Get-ChildItem $Root -Filter *.xlsx | Where-Object { $_.Name -notlike '~$*' }
    if (-not $mine) { Copy-Item "$src\*.xlsx" $Root -ErrorAction SilentlyContinue }

    Write-Host "Updated to version $remote. Your results and school list were not touched."
} finally {
    Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
}
