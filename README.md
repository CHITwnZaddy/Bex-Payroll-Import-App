# BEX Payroll Import App

Small Windows desktop app for creating Vista payroll import CSV files from BEX payroll source reports.

Phase 1 keeps the existing Payroll Upload Worksheet formulas as the calculation source. The app will copy the template for each run, load the Time Detail Report and optional Expense Transaction file, recalculate through desktop Excel, and export the `PU` tab as a headerless CSV.

See the design spec:

[/docs/superpowers/specs/2026-05-29-bex-payroll-import-app-design.md](docs/superpowers/specs/2026-05-29-bex-payroll-import-app-design.md)

Real payroll workbooks and generated payroll CSV files should stay out of git.
