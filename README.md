# Sales Tracker App

This is a simple Streamlit app built from your workbook model so users can work with a clean dashboard instead of raw sheets.

## What it does

- Reads data from `Get.1a.Sales Tracker.polished.xlsx`
- Shows KPI cards for approaches, meetings, proposals, and secured value
- Provides filters for lead staff, prospect status, and date range
- Displays interactive charts for:
  - work secured by staff
  - monthly secured value trend
  - prospect status mix
- Includes pipeline, team, and COI tables
- Lets users download filtered pipeline data as CSV
- Creates a timestamped backup copy before workbook save operations

## Run locally

1. Open a terminal in this folder.
2. Install dependencies:

```powershell
pip install -r requirements.txt
```

3. Start the app:

```powershell
streamlit run app.py
```

4. Open the local URL shown in terminal (usually http://localhost:8501).

## One-click launch

- Double-click `start-app.bat`
- Or run `run-app.ps1` in PowerShell

The launcher creates a local virtual environment, installs requirements, and starts the app.

## Build a Windows executable

1. Install Python 3.11+ on the build machine.
2. Put the workbook in this folder as:

`Get.1a.Sales Tracker.polished.xlsx`

3. Double-click `build-exe.bat` (or run `build-exe.ps1`).
4. After build, open:

`dist\\SalesTrackerApp\\SalesTrackerApp.exe`

### Share with your team

- Zip and share the full folder `dist\\SalesTrackerApp`
- Team members run `SalesTrackerApp.exe` directly (no Python install required)

## Workbook path

The app defaults to:

E:\Visual Code Projects\Get.1a.Sales Tracker.polished.xlsx

You can change the workbook path from the sidebar if needed.

The app now validates the path before loading. It must point to an existing `.xlsx` or `.xlsm` file.

If workbook values rely on unsupported Excel formulas in the Stats to Date sheet, the app shows `Unsupported formula` instead of leaving the cell blank.

## Backups

Before saving Pipeline, COI, or list changes back to the workbook, the app creates a timestamped backup copy in:

`backups\`

## Pipeline import template

The Pipeline import template download includes:

- matching Pipeline headers
- dropdown validations for supported list fields
- an `Instructions` sheet
- a highlighted sample row on the import sheet

## Manager access

Firm Manager sign-in is configured from either Streamlit secrets or environment variables.

The app can also save a local manager password directly from the sidebar. That value is stored in:

`app_config.json`

Supported names:

- `manager_password` in Streamlit secrets
- `MANAGER_PASSWORD` in Streamlit secrets
- `SALES_TRACKER_MANAGER_PASSWORD` environment variable
- `MANAGER_PASSWORD` environment variable

Example PowerShell session before launch:

```powershell
$env:SALES_TRACKER_MANAGER_PASSWORD = "your-secure-password"
streamlit run app.py
```

## Contributing

1. Create a feature branch from `main`.
2. Install dependencies with `pip install -r requirements.txt`.
3. Run quick checks before opening a pull request:

```powershell
python test_imports.py
python -m py_compile app.py launcher.py test_imports.py
```

4. Open a pull request to `main` with a short summary of changes and testing.

## GitHub CI

This repository includes a GitHub Actions workflow at `.github/workflows/ci.yml`.

It runs on pushes and pull requests to `main` and performs:

- dependency installation from `requirements.txt`
- import validation via `test_imports.py`
- Python syntax compilation checks

## Branch Protection (Recommended)

Set branch protection for `main` in GitHub:

1. Open repository settings.
2. Go to Branches -> Add branch protection rule.
3. Choose `main`.
4. Enable:
  - Require a pull request before merging
  - Require status checks to pass before merging (`CI`)
  - Optional: require linear history and prevent force pushes
