$ErrorActionPreference = 'Stop'

Set-Location $PSScriptRoot

function Get-PythonCommand {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        return 'py'
    }
    if (Get-Command python -ErrorAction SilentlyContinue) {
        return 'python'
    }
    throw 'Python is not installed. Install Python 3.11+ first.'
}

$pythonCmd = Get-PythonCommand

$workbookName = 'Get.1a.Sales Tracker.polished.xlsx'
$localWorkbook = Join-Path $PSScriptRoot $workbookName
$parentWorkbook = Join-Path (Split-Path $PSScriptRoot -Parent) $workbookName

if (-not (Test-Path $localWorkbook)) {
    if (Test-Path $parentWorkbook) {
        Copy-Item $parentWorkbook $localWorkbook -Force
    } else {
        throw "Workbook not found. Expected file: $workbookName"
    }
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
    throw 'Virtual environment could not be created.'
}

& $venvPython -m pip install --upgrade pip
& $venvPython -m pip install -r requirements.txt
& $venvPython -m pip install pyinstaller==6.12.0

$distDir = Join-Path $PSScriptRoot 'dist'
$buildDir = Join-Path $PSScriptRoot 'build'
if (Test-Path $distDir) { Remove-Item $distDir -Recurse -Force }
if (Test-Path $buildDir) { Remove-Item $buildDir -Recurse -Force }

& $venvPython -m PyInstaller `
    --noconfirm `
    --onedir `
    --name SalesTrackerApp `
    --collect-all streamlit `
    --collect-all plotly `
    --collect-all pandas `
    --collect-all openpyxl `
    --hidden-import pyarrow `
    --add-data "app.py;." `
    --add-data "Get.1a.Sales Tracker.polished.xlsx;." `
    launcher.py

Write-Host ''
Write-Host 'Build complete.' -ForegroundColor Green
Write-Host 'Executable folder:' (Join-Path $PSScriptRoot 'dist\SalesTrackerApp')
Write-Host 'Run:' (Join-Path $PSScriptRoot 'dist\SalesTrackerApp\SalesTrackerApp.exe')
