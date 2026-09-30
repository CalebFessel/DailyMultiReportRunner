<#
.SYNOPSIS
    Generates the daily report workbooks on Windows and opens the output folder.

.DESCRIPTION
    Finds a usable Python, creates a local virtual environment, installs the
    dependencies, and runs the API-backed report bundle. No email is attempted
    unless SMTP credentials and recipients are configured, so by default this
    just produces files for you to send yourself.

    Credentials come from a .env file next to this script (copy .env.example
    to .env and fill in TS_API_BASE_URL, TS_API_KEY and TS_API_SECRET).

.EXAMPLE
    .\run_reports.ps1
    Yesterday's reports, zipped, folder opened.

.EXAMPLE
    .\run_reports.ps1 -Date 2026-08-17 -NoZip

.NOTES
    If PowerShell refuses to run this file:
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#>
[CmdletBinding()]
param(
    [string]$Date,
    [switch]$NoZip,
    [switch]$NoOpen
)

$ErrorActionPreference = 'Stop'
Set-Location -Path $PSScriptRoot

function Find-Python {
    # The Microsoft Store alias answers to "python" without being Python, so
    # every candidate is validated by actually asking for a version.
    foreach ($candidate in @(
        @{ Exe = 'py';      Args = @('-3', '--version') },
        @{ Exe = 'python3'; Args = @('--version') },
        @{ Exe = 'python';  Args = @('--version') }
    )) {
        if (-not (Get-Command $candidate.Exe -ErrorAction SilentlyContinue)) { continue }
        try { $version = & $candidate.Exe @($candidate.Args) 2>&1 } catch { continue }
        if ($LASTEXITCODE -eq 0 -and "$version" -match 'Python (\d+)\.(\d+)') {
            if ([int]$Matches[1] -gt 3 -or ([int]$Matches[1] -eq 3 -and [int]$Matches[2] -ge 10)) {
                $prefix = @()
                if ($candidate.Exe -eq 'py') { $prefix = @('-3') }
                return [pscustomobject]@{ Exe = $candidate.Exe; Prefix = $prefix; Version = "$version".Trim() }
            }
            Write-Warning "$($candidate.Exe) is $version; 3.10 or newer is required."
        }
    }
    return $null
}

$python = Find-Python
if (-not $python) {
    Write-Host ''
    Write-Host 'No usable Python 3.10+ was found.' -ForegroundColor Red
    Write-Host ''
    Write-Host 'Install Python 3.12 from https://www.python.org/downloads/windows/'
    Write-Host 'and tick "Add python.exe to PATH" during setup, then run this again.'
    Write-Host ''
    Write-Host '"Python was not found; run without arguments to install from the Microsoft'
    Write-Host 'Store" means Windows'' app-execution alias answered instead of a real'
    Write-Host 'interpreter -- installing from python.org replaces it.'
    Write-Host ''
    exit 1
}
Write-Host "Using $($python.Version)" -ForegroundColor Green

# --- environment ---
$venv = Join-Path $PSScriptRoot '.venv'
$venvPython = Join-Path $venv 'Scripts\python.exe'
if (-not (Test-Path $venvPython)) {
    Write-Host 'Creating virtual environment in .venv (one time) ...'
    & $python.Exe @($python.Prefix + @('-m', 'venv', $venv))
    if ($LASTEXITCODE -ne 0) { throw 'Failed to create the virtual environment.' }
}

Write-Host 'Checking dependencies ...'
& $venvPython -m pip install --quiet --upgrade pip
& $venvPython -m pip install --quiet requests python-dotenv pandas openpyxl
if ($LASTEXITCODE -ne 0) { throw 'Failed to install dependencies.' }

# --- credentials ---
if (-not (Test-Path (Join-Path $PSScriptRoot '.env'))) {
    if (-not $env:TS_API_KEY) {
        Write-Host ''
        Write-Host 'No .env file and no TS_API_KEY in the environment.' -ForegroundColor Yellow
        Write-Host ''
        Write-Host '  Copy-Item .env.example .env'
        Write-Host '  notepad .env'
        Write-Host ''
        Write-Host 'Fill in TS_API_BASE_URL, TS_API_KEY and TS_API_SECRET, then run this again.'
        Write-Host ''
        exit 1
    }
}

# --- run ---
$reportArgs = @()
if ($Date) { $reportArgs += $Date }
if (-not $NoZip) { $reportArgs += '--zip' }

Write-Host ''
Write-Host 'Generating reports ...' -ForegroundColor Cyan
Write-Host ''
& $venvPython (Join-Path $PSScriptRoot 'daily_report_runner_api.py') @reportArgs
$code = $LASTEXITCODE

$outputDir = $env:OUTPUT_DIR
if (-not $outputDir) { $outputDir = 'Reports' }
$outputPath = Join-Path $PSScriptRoot $outputDir

if ($code -eq 0) {
    Write-Host ''
    Write-Host "Done. Workbooks are in $outputPath" -ForegroundColor Green
    if (-not $NoOpen -and (Test-Path $outputPath)) { Start-Process explorer.exe $outputPath }
}
else {
    Write-Host ''
    Write-Host "The run reported errors (exit $code). Check the log under $outputPath\logs." -ForegroundColor Yellow
}
exit $code
