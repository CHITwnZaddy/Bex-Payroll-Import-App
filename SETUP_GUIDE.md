# Setup guide

## Target machine

This app is built for one Windows user with:

- Windows
- local Microsoft Excel installed
- Python installed
- `openpyxl`
- `pywin32`

## Install dependencies

Open PowerShell in the project folder:

```powershell
py -m pip install -e ".[dev,windows]"
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

Before importing into Vista, open the generated CSV and confirm:

- no header row
- one batch code on every row
- expected employee codes
- expense rows have `Hours=0`
- expense rows use `Department=DLPTO`
- expense rows use `Pay type=EXP REIM`
- expense rows use `Cost Type=L`
