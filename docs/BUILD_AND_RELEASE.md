# Build instructions (Excel 2019, Windows 11)

This repository currently publishes VBA source modules rather than a compiled or runtime-tested XLSM release.

## Assemble a workbook
1. In Excel 2019 Desktop, create a blank workbook and save it as ExcelDataReconciliation.xlsm.
2. Press Alt+F11. Import these files with File > Import File:
   - src/modReconciliation.bas
   - src/modDashboard.bas
   - src/modAdvancedReconciliation.bas
   - src/modReportPresentation.bas
   - src/modUnifiedDashboard.bas
3. Run Debug > Compile VBAProject.
4. Execute BuildUnifiedDashboard via Alt+F8. This creates a worksheet named Dashboard (or renames an older Reconciliation v0.4 sheet).
5. Use the composite CSV demo and compare results with tests/REPORTING_TESTS.md.
6. Save the workbook, then reopen and test again before sharing it with clients.

If upgrading an existing workbook, replace the older modUnifiedDashboard module with the latest version; do not import a second module with the same name.

## Release status
- v0.4: basic operation user-confirmed on Windows 11 / Excel 2019.
- v0.5: basic functionality confirmed by project owner in Excel 2019.
- v0.6: clean interface upgrade source published; needs Excel 2019 testing before shipping.
- Other Office versions and Mac are not validated.
- Excel for the web does not execute VBA.
- An XLSM release file will be added only after testing and review.

## Clean delivery
Use only fictional demo datasets. Remove saved local paths and Run History entries before sharing a workbook publicly. Keep macro security policies enabled. Check XLSX and PDF exports manually.

## Clean v0.6 release (recommended)

For the unified workflow, only three source modules are needed: `modAdvancedReconciliation.bas`, `modReportPresentation.bas`, and `modUnifiedDashboard.bas`. Historical `modReconciliation.bas` and `modDashboard.bas` may be omitted in a newly built workbook. Run `BuildUnifiedDashboard` and verify that its only application worksheets are `Dashboard` and `Run History` after the first comparison. Generated report workbooks have four sheets, including KPI Dashboard.

For upgrading an existing file, preserve the previous tested `.xlsm` as a backup. Follow `tests/V06_ACCEPTANCE.md`; do not remove existing modules blindly.
