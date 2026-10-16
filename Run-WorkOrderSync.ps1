<#
.SYNOPSIS
  Runs TraumaSoft_Samsara_WorkOrder_Sync.py with whatever Python is available on this box.

.DESCRIPTION
  Works in Windows PowerShell 5.1 and PowerShell 7. Looks for Python in this order:
    1. $env:PYTHON_EXE (set it in .env or the Task Scheduler action if you want a fixed path)
    2. .venv\Scripts\python.exe next to this script
    3. the "py" launcher (py -3)
    4. python / python3 on PATH
    5. common per-user / system install folders
  Passes every argument straight through to the sync script.

.EXAMPLE
  .\Run-WorkOrderSync.ps1 --discover
  .\Run-WorkOrderSync.ps1 --dry-run --no-email
  .\Run-WorkOrderSync.ps1 -Install          # pip install -r requirements.txt first
#>
[CmdletBinding()]
param(
    [switch]$Install,
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$SyncArgs
)

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

function Find-Python {
    if ($env:PYTHON_EXE -and (Test-Path $env:PYTHON_EXE)) { return @($env:PYTHON_EXE) }
    $venv = Join-Path $here ".venv\Scripts\python.exe"
    if (Test-Path $venv) { return @($venv) }
    if (Get-Command py -ErrorAction SilentlyContinue) { return @("py", "-3") }
    foreach ($name in @("python", "python3")) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        # Skip the Microsoft Store alias stub, which only prints "Python was not found".
        if ($cmd -and $cmd.Source -notmatch "WindowsApps") { return @($cmd.Source) }
    }
    $candidates = @(
        "$env:LOCALAPPDATA\Programs\Python\Python3*\python.exe",
        "C:\Python3*\python.exe",
        "C:\Program Files\Python3*\python.exe",
        "$env:ProgramFiles\Python3*\python.exe"
    )
    foreach ($pattern in $candidates) {
        $hit = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue | Sort-Object FullName -Descending | Select-Object -First 1
        if ($hit) { return @($hit.FullName) }
    }
    throw "No Python found. Install Python 3.10+ from python.org (tick 'Add to PATH'), or set PYTHON_EXE to the full path of python.exe."
}

$py = Find-Python
$exe = $py[0]
$pre = @()
if ($py.Count -gt 1) { $pre = $py[1..($py.Count - 1)] }
Write-Host "Using Python: $($py -join ' ')"

if ($Install) {
    & $exe @pre -m pip install -r (Join-Path $here "requirements.txt")
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

& $exe @pre (Join-Path $here "TraumaSoft_Samsara_WorkOrder_Sync.py") @SyncArgs
exit $LASTEXITCODE
