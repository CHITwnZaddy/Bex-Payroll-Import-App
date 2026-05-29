# BEX payroll import app design

Date: 2026-05-29

## Purpose

Build a small Windows desktop app that turns BEX payroll source files into a Vista payroll import CSV.

Phase 1 keeps the existing Payroll Upload Worksheet as the calculation source. The workbook already contains lookup logic, formulas, and field mappings that the business trusts today. The app wraps that process so the user stops copying sheets by hand and stops saving the wrong tab or file type.

## Current manual process

The user currently:

1. Opens the Time Detail Report.
2. Copies it into the `TDR` tab of the Payroll Upload Worksheet.
3. Saves the workbook so formulas refresh.
4. Activates the `PU` tab.
5. Saves the workbook as CSV.
6. Imports that CSV into Vista.

Expense transactions are handled inconsistently today. Phase 1 pulls them into the same run so the user has one repeatable path.

## Phase 1 app shape

Use Python with Tkinter for a tiny Windows desktop app.

Use `openpyxl` for workbook inspection and row writing. Use desktop Excel through `pywin32` for recalculation and CSV export because `openpyxl` preserves formulas but does not calculate them.

Python and `openpyxl` are already installed on the target machine. The setup guide must include a `pywin32` check and install command:

```powershell
py -m pip install pywin32
```

The normal user should run the app, choose files, and click generate. PowerShell belongs in setup and troubleshooting only.

## First-run setup

On first launch, the app asks the user to select the Payroll Upload Worksheet template.

The app saves that template path locally. Future runs reuse it. The main screen includes a settings action to change the template path.

If the saved template path breaks, the app stops and asks the user to select the template again.

The app never edits the saved template directly. Each run copies the template into the output folder and works from that copied workbook.

## Normal run workflow

The app has one main screen with:

- saved template status
- Time Detail Report picker
- Expense Transaction picker
- checkbox for confirming there is no Expense Transaction file
- Generate button
- status area with the output folder and validation result

The Expense Transaction file is a soft requirement. The user must either select the file or check a confirmation box that says they intentionally have no Expense Transaction file for this payroll import.

The Generate button stays disabled until the Time Detail Report exists and the expense decision is made.

## Output folder

Outputs are created beside the selected Time Detail Report.

Folder pattern:

```text
<TDR folder>\Payroll Import Outputs\YYYY-MM-DD\
```

Example:

```text
...\BEX Payroll - #05\Payroll Import Outputs\2026-05-29\
```

Each run writes files using the generated batch code and timestamp so repeated same-day backfill runs do not collide.

## Batch code

Every generated import gets a fresh batch code.

Format:

```text
CDMMDDHHMMSS
```

Example:

```text
CD0529143708
```

Length: 12 characters. That fits the Vista batch code limit discussed during design.

The app replaces or sets the batch code on every generated row. It does this for rows from the Time Detail Report and rows normalized from Expense Transaction.

The output filename can include the year even though the batch code does not:

```text
PayrollImport_CD0529143708_2026-05-29_14-37-08.csv
```

## Source files

### Payroll Upload Worksheet template

The template has these relevant sheets:

- `PU`: final import output
- `TDR`: source data and helper formulas
- `Key`: employee lookup data

The `HR` sheet is ignored for phase 1 unless existing formulas reference it.

### Time Detail Report

The app reads the Time Detail Report and writes its rows into the copied template's `TDR` tab.

The source columns needed for the final import are:

- Batch code
- Employee code
- Department
- Pay type
- Hours
- Job
- Phase
- Cost Type
- Date

Some of these are calculated by the template formulas after the Time Detail Report data lands on the `TDR` tab.

### Expense Transaction

The app reads Expense Transaction rows and appends normalized rows to the copied template's `TDR` tab.

Expense row mapping:

| Output concept | Expense source or value |
| --- | --- |
| Date | `Expense Date` |
| Employee | parse `Employee` into first and last name, then resolve through `Key` |
| Hours | `0` |
| Pay type | `EXP REIM` |
| Department | `DLPTO` |
| Job | blank |
| Phase | blank |
| Cost Type | `L` |
| Dollars | `Paid Amount` |
| Batch code | generated `CDMMDDHHMMSS` |

The final CSV includes the dollar column in phase 1. If Vista rejects that import shape, remove that column in a follow-up change after the first failed import proves the issue.

## Workbook behavior

The copied workbook is the audit workbook for the run.

The app writes source rows to `TDR`, extends helper formulas when the source rows exceed the template's existing formula range, opens the workbook in Excel, recalculates, saves, and exports the `PU` output to CSV.

The final CSV drops headers.

The app should preserve the workbook's formulas, tabs, and lookup behavior. Phase 1 avoids rebuilding the formula logic in Python.

## Validation rules

Validation has 2 levels: hard stops and warnings.

Hard stops block final CSV creation. The app still creates the output folder and writes a validation file. It writes an audit workbook when it can.

Hard stops:

- missing Time Detail Report
- missing or invalid template
- missing required source columns
- selected source file cannot be opened
- selected source file has no usable data rows
- unmatched employee from Expense Transaction
- TDR-derived lookup fields return `NL` for required fields
- required output fields are blank after Excel recalculation
- Excel cannot open, recalculate, save, or export
- final `PU` output has no valid rows

Warnings allow final CSV creation.

Warnings:

- Expense Transaction file intentionally skipped
- dollar column included in final CSV
- expense rows use `Hours=0`
- expense rows have blank Job and Phase
- output date folder already exists
- row count is unusually low or unusually high

Each hard stop has a plain reason and a suggested fix.

Example:

```text
Row 12: Expense employee "DOE, JANE" could not be matched to the Key tab.
Possible fix: confirm the employee name in the Expense Transaction file or add/correct the employee in the template Key tab, then rerun.
```

The app should show a short message and offer to open the validation file. Stack traces should go into a technical log, not the normal user message.

## Output files

Successful run:

- `PayrollImport_<batch>_<timestamp>.csv`
- `PayrollUpload_Audit_<batch>_<timestamp>.xlsx`
- `ValidationReport_<batch>_<timestamp>.csv`

Hard-stop run:

- `PayrollUpload_Audit_<batch>_<timestamp>.xlsx`, when available
- `ValidationErrors_<batch>_<timestamp>.csv`
- technical log, if the app crashes or Excel automation fails

The final payroll import CSV should contain data rows only. Headers are removed.

## Testing plan

Test first with the 3 provided sample files:

1. Payroll Upload Worksheet template.
2. Time Detail Report.
3. Expense Transaction file.

Happy-path test:

1. Select the saved template.
2. Select the Time Detail Report.
3. Select the Expense Transaction file.
4. Generate output.
5. Confirm the output folder exists.
6. Confirm the audit workbook opens in Excel.
7. Confirm Excel recalculated the `PU` tab.
8. Confirm the final CSV has no headers.
9. Confirm all rows use one generated batch code.
10. Confirm expense rows use the agreed defaults.

Failure tests:

1. Leave Expense Transaction blank and leave the confirmation unchecked. Generate should stay disabled.
2. Check the no-expense confirmation and generate. The app should warn, then create a valid CSV from TDR rows only.
3. Change an expense employee name so it cannot match `Key`. The app should hard stop and write a validation error.
4. Remove a required column from a copy of the Time Detail Report. The app should hard stop and name the missing column.
5. Temporarily break the saved template path. The app should ask for the template again.

## Handoff docs

Create 2 docs with the app:

- `README.md`: what the tool does, required files, output files, assumptions, support notes.
- `SETUP_GUIDE.md`: install and check Python packages, first launch, template selection, test run, common errors.

Screenshots can be added after the UI exists.

## Definition of done

Phase 1 is done when:

- the coworker can run a tiny Windows app for normal use
- the template path saves after first setup
- the app refuses to generate until the Time Detail Report and expense decision are present
- every run gets a unique `CDMMDDHHMMSS` batch code
- expense rows are included with the agreed defaults
- Excel recalculates before CSV export
- the CSV has no headers
- hard-stop failures write a useful review file
- the saved Payroll Upload Worksheet template remains untouched

## Known technical debt

Phase 1 depends on desktop Excel and the current Payroll Upload Worksheet formulas.

The first version should keep that dependency because the workbook is already the business source of truth. A later version should replace the workbook formulas with explicit code after the import format and expense behavior are proven.
