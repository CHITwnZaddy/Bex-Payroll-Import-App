# BEX Payroll Import App

Small Windows desktop app for creating Vista payroll import CSV files from BEX payroll source reports.

## What it does

The app takes:

- Payroll Upload Worksheet template
- Time Detail Report
- Expense Transaction file, or a user confirmation that no expense file exists

It creates:

- payroll import CSV with no headers
- audit workbook
- validation report
- validation error file when a run is blocked

## Safety rules

The app copies the Payroll Upload Worksheet template for every run. It does not edit the saved template.

Real payroll workbooks and generated CSV files should stay out of git.

## Output folder

Outputs are created beside the selected Time Detail Report:

```text
<TDR folder>\Payroll Import Outputs\YYYY-MM-DD\
```

## Batch code

Every run gets a fresh batch code:

```text
CDMMDDHHMMSS
```

Example:

```text
CD0529143708
```

## Development

Install:

```bash
python3 -m pip install -e ".[dev]"
```

Run tests:

```bash
python3 -m pytest -q
```

Windows install with Excel automation support:

```powershell
py -m pip install -e ".[dev,windows]"
```
