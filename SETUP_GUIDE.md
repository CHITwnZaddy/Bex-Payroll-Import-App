# Setup guide

## Target machine

This app is built for one Windows user with:

- Windows
- local Microsoft Excel installed
- Python 3.10+ installed
- `openpyxl`
- `pywin32`

## Install dependencies

Open PowerShell in the project folder:

```powershell
py --version
```

```powershell
py -m pip install -e ".[windows]"
```

If `pywin32` needs to be installed separately:

```powershell
py -m pip install pywin32
```

## Start the app

```powershell
py -m bex_payroll_import
```

## First run

1. Select the Payroll Upload Worksheet template.
2. Select the Time Detail Report.
3. Select the Expense Transaction file, or check the box confirming there is no Expense Transaction file.
4. Click Generate Payroll Import.
5. Review the output folder.

## Success output

The app creates:

- `PayrollImport_<batch>_<timestamp>.csv`
- `PayrollUpload_Audit_<batch>_<timestamp>.xlsx`
- `ValidationReport_<batch>_<timestamp>.csv`

## Blocked run output

The app creates:

- `ValidationErrors_<batch>_<timestamp>.csv`
- audit workbook when the app reached workbook creation

Open the validation error file and correct the source file or template, then rerun.

## First live test

Use a known payroll period first.

Before importing into Vista, inspect the generated CSV in Notepad without saving it. Confirm:

- no header row
- one batch code on every row
- expected employee codes
- expense rows have blank Hours, Job, Phase, and Cost Type
- expense rows use `Department=1`
- expense rows use `Pay type=EXP REIMB`
- expense rows use `Check Date` for Date
- expense rows use `Paid Amount` for Dollars

If you open the CSV in Excel, close it without saving. Import the original generated CSV file into Vista.

## Optional packaged app

Build the `.exe` on Windows:

```powershell
py -m PyInstaller --onefile --windowed --name BEXPayrollImport src\bex_payroll_import\__main__.py
```

Run:

```powershell
.\dist\BEXPayrollImport.exe
```

Give the user the `.exe` only after a successful test against a known payroll period.
