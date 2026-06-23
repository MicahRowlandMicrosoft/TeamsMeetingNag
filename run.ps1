# Bootstrap script: create venv, install deps, then run the nag.
# Usage:  .\run.ps1            (run)
#         .\run.ps1 -LoginOnly  (sign in and cache token)
#         .\run.ps1 -Reinstall  (force reinstall of dependencies)
[CmdletBinding()]
param(
    [switch]$LoginOnly,
    [switch]$Reinstall
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$venv = Join-Path $root ".venv"
$python = Join-Path $venv "Scripts\python.exe"
$stamp = Join-Path $venv "deps.installed"

# Pick the best available Python. cryptography wheels exist for cp313 win_arm64
# but not all newer/older combos - prefer 3.13 if present.
function Get-PythonLauncher {
    $candidates = @("-3.13", "-3.12", "-3.11", "-3")
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

Write-Host ""
Write-Host "Starting Teams Meeting Nag..." -ForegroundColor Green
Write-Host "On first run a browser will open for Microsoft sign-in." -ForegroundColor Yellow
Write-Host "Press Ctrl+C to stop." -ForegroundColor DarkGray
Write-Host ""

& $python @pyArgs
