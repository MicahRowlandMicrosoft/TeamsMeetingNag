# Bootstrap script: create venv, install deps, then run the nag.
# Usage:  .\run.ps1            (run)
#         .\run.ps1 -Config     (open the settings UI)
#         .\run.ps1 -LoginOnly  (sign in and cache token)
#         .\run.ps1 -Reinstall  (force reinstall of dependencies)
[CmdletBinding()]
param(
    [switch]$LoginOnly,
    [switch]$Reinstall,
    [switch]$Config
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$venv = Join-Path $root ".venv"
$python = Join-Path $venv "Scripts\python.exe"
$stamp = Join-Path $venv "deps.installed"

# Pick the best available Python. We prefer 3.14 because it bundles tkinter
# on win_arm64 (the settings UI needs it); 3.13-arm64 does not. cryptography
# wheels are not required by our deps, so older minor versions also work.
function Get-PythonLauncher {
    $candidates = @("-3.14", "-3.13", "-3.12", "-3.11", "-3")
    foreach ($c in $candidates) {
        $v = & py $c -c "import sys;print(sys.version)" 2>$null
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Using Python $c ($($v.Split()[0]))" -ForegroundColor DarkGray
            return $c
        }
    }
    throw "No usable Python found. Install Python 3.11+ from https://python.org or the Microsoft Store."
}

if (-not (Test-Path $python)) {
    $pyArg = Get-PythonLauncher
    Write-Host "Creating virtual environment in .venv ..." -ForegroundColor Cyan
    & py $pyArg -m venv $venv
}

# The settings UI needs tkinter. Win-ARM64 Python 3.13 ships without it; 3.14
# does. If -Config was requested and the current venv can't import tkinter,
# rebuild the venv on a Python that has it.
if ($Config) {
    & $python -c "import tkinter" 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Current venv Python lacks tkinter; rebuilding .venv on a Python that has it..." -ForegroundColor Yellow
        Remove-Item -Recurse -Force $venv
        $pyArg = Get-PythonLauncher
        & py $pyArg -c "import tkinter" 2>$null
        if ($LASTEXITCODE -ne 0) {
            throw "No installed Python has tkinter. Install Python 3.14 from https://python.org and re-run."
        }
        & py $pyArg -m venv $venv
        # Force a dependency reinstall on the new interpreter.
        Remove-Item -Force -ErrorAction SilentlyContinue $stamp
    }
}

if ($Reinstall -or -not (Test-Path $stamp)) {
    Write-Host "Installing dependencies (first run can take a minute)..." -ForegroundColor Cyan
    & $python -m pip install --upgrade pip
    if ($LASTEXITCODE -ne 0) { throw "pip upgrade failed." }
    & $python -m pip install -r (Join-Path $root "requirements.txt")
    if ($LASTEXITCODE -ne 0) { throw "Dependency install failed." }
    Set-Content -Path $stamp -Value (Get-Date -Format o)
} else {
    Write-Host "Dependencies already installed (delete .venv\deps.installed or pass -Reinstall to refresh)." -ForegroundColor DarkGray
}

$pyArgs = @((Join-Path $root "nag.py"))
if ($LoginOnly) { $pyArgs += "--login-only" }

if ($Config) {
    Write-Host ""
    Write-Host "Opening configuration UI..." -ForegroundColor Green
    & $python (Join-Path $root "config_ui.py")
    exit $LASTEXITCODE
}

Write-Host ""
Write-Host "Starting Teams Meeting Nag..." -ForegroundColor Green
Write-Host "On first run a browser will open for Microsoft sign-in." -ForegroundColor Yellow
Write-Host "Press Ctrl+C to stop." -ForegroundColor DarkGray
Write-Host ""

& $python @pyArgs
