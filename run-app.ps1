$ErrorActionPreference = 'Stop'

Set-Location $PSScriptRoot

$pythonCmd = $null
if (Get-Command py -ErrorAction SilentlyContinue) {
    $pythonCmd = 'py'
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $pythonCmd = 'python'
}

if (-not $pythonCmd) {
    Write-Host 'Python is not installed on this machine.' -ForegroundColor Yellow
    Write-Host 'Install Python 3.11+ and re-run this script.' -ForegroundColor Yellow
    exit 1
}

if (-not (Test-Path '.venv')) {
    if ($pythonCmd -eq 'py') {
        py -3 -m venv .venv
    } else {
        python -m venv .venv
    }
}

$venvPython = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (-not (Test-Path $venvPython)) {
    Write-Host 'Virtual environment creation failed.' -ForegroundColor Red
    exit 1
}

& $venvPython -m pip install --upgrade pip
& $venvPython -m pip install -r requirements.txt
& $venvPython -m streamlit run app.py
