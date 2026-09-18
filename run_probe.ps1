<#
.SYNOPSIS
    Runs the Traumasoft ThirdParty API probe on Windows.

.DESCRIPTION
    Finds a usable Python, creates a local virtual environment, installs the
    dependencies, and runs probe_traumasoft_api.py. Credentials are read from
    a .env file in this folder (copy .env.example to .env and fill it in).

    The probe is read-only: it issues GET requests only.

.EXAMPLE
    .\run_probe.ps1
    .\run_probe.ps1 -Date 2026-08-17

.NOTES
    If PowerShell refuses to run this file, unblock it for the current session:
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#>
[CmdletBinding()]
param(
    [string]$Date = (Get-Date).AddDays(-1).ToString('yyyy-MM-dd'),
    [string]$OutDir = 'api_probe'
)

$ErrorActionPreference = 'Stop'
Set-Location -Path $PSScriptRoot

function Find-Python {
    # The Microsoft Store alias stub answers to "python" but is not Python, so
    # candidates are validated by actually asking for a version.
    foreach ($candidate in @(
        @{ Exe = 'py';      Args = @('-3', '--version') },
        @{ Exe = 'python3'; Args = @('--version') },
        @{ Exe = 'python';  Args = @('--version') }
    )) {
        $cmd = Get-Command $candidate.Exe -ErrorAction SilentlyContinue
        if (-not $cmd) { continue }
        try {
            $version = & $candidate.Exe @($candidate.Args) 2>&1
        } catch {
            continue
        }
        if ($LASTEXITCODE -eq 0 -and "$version" -match 'Python (\d+)\.(\d+)') {
            $major = [int]$Matches[1]
            $minor = [int]$Matches[2]
            if ($major -gt 3 -or ($major -eq 3 -and $minor -ge 10)) {
                $prefix = @()
                if ($candidate.Exe -eq 'py') { $prefix = @('-3') }
                return [pscustomobject]@{
                    Exe     = $candidate.Exe
                    Prefix  = $prefix
                    Version = "$version".Trim()
                }
            }
            Write-Warning "$($candidate.Exe) is $version; the report runner needs 3.10 or newer."
        }
    }
    return $null
}

$python = Find-Python
if (-not $python) {
    Write-Host ''
    Write-Host 'No usable Python 3.10+ was found on this machine.' -ForegroundColor Red
    Write-Host ''
    Write-Host 'The message "Python was not found; run without arguments to install from the'
    Write-Host 'Microsoft Store" means Windows'' app-execution alias answered instead of a real'
    Write-Host 'interpreter. Two ways forward:'
    Write-Host ''
    Write-Host '  1. Use the Python that already runs the daily report. Find it with:'
    Write-Host '       Get-Command python, python3, py -ErrorAction SilentlyContinue'
    Write-Host '       Get-ChildItem C:\Python*, "$env:LOCALAPPDATA\Programs\Python" -ErrorAction SilentlyContinue'
    Write-Host ''
    Write-Host '  2. Install Python 3.12 from https://www.python.org/downloads/windows/'
    Write-Host '     and tick "Add python.exe to PATH" during setup.'
    Write-Host ''
    exit 1
}
Write-Host "Using $($python.Version) via '$($python.Exe)'" -ForegroundColor Green

# --- virtual environment ---
$venv = Join-Path $PSScriptRoot '.venv'
$venvPython = Join-Path $venv 'Scripts\python.exe'
if (-not (Test-Path $venvPython)) {
    Write-Host 'Creating virtual environment in .venv ...'
    & $python.Exe @($python.Prefix + @('-m', 'venv', $venv))
    if ($LASTEXITCODE -ne 0) { throw 'Failed to create the virtual environment.' }
}

Write-Host 'Installing dependencies ...'
& $venvPython -m pip install --quiet --upgrade pip
& $venvPython -m pip install --quiet requests python-dotenv
if ($LASTEXITCODE -ne 0) { throw 'Failed to install dependencies.' }

# --- credentials ---
$envFile = Join-Path $PSScriptRoot '.env'
if (-not (Test-Path $envFile)) {
    Write-Host ''
    Write-Host 'No .env file found.' -ForegroundColor Yellow
    Write-Host 'Copy .env.example to .env and fill in TS_API_BASE_URL, TS_API_KEY and TS_API_SECRET.'
    Write-Host ''
    Write-Host '  Copy-Item .env.example .env'
    Write-Host '  notepad .env'
    Write-Host ''
    Write-Host 'To set them for this session instead (PowerShell uses $env:, not export):'
    Write-Host '  $env:TS_API_BASE_URL = "https://lynxems.traumasoft.com"'
    Write-Host '  $env:TS_API_KEY      = "..."'
    Write-Host '  $env:TS_API_SECRET   = "..."'
    Write-Host ''
    if (-not $env:TS_API_KEY -or -not $env:TS_API_SECRET) { exit 1 }
    Write-Host 'Falling back to the environment variables already set in this session.'
}

# --- run ---
Write-Host ''
Write-Host "Probing the API for $Date ..." -ForegroundColor Cyan
& $venvPython (Join-Path $PSScriptRoot 'probe_traumasoft_api.py') $Date --out $OutDir
$code = $LASTEXITCODE

if ($code -eq 0) {
    $findings = Join-Path $OutDir 'PROBE_FINDINGS.md'
    Write-Host ''
    Write-Host "Findings written to $findings" -ForegroundColor Green
    Write-Host 'Raw samples are alongside it as *_sample.json.'
}
exit $code
