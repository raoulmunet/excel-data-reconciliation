# Clean XLSM build — Excel 2019 on Windows 11

Current production source files (and **only these three**):
- `src/modAdvancedReconciliation.bas`
- `src/modReportPresentation.bas`
- `src/modUnifiedDashboard.bas`

Legacy modules were removed from the main branch and preserved in `archive/pre-clean-build-2026-10-08`.

## Automated build (PowerShell + installed Excel)

1. Download/extract the **whole repository**.
2. In Excel: File > Options > Trust Center > Trust Center Settings > Macro Settings > enable **Trust access to the VBA project object model** temporarily. Do not disable other macro protections.
3. Close Excel. Open **Windows PowerShell** in the repository root.
4. Run `powershell.exe -NoProfile -File .\build\Create-CleanWorkbook.ps1`.
5. The builder refuses to overwrite an existing `ExcelDataReconciliation.xlsm`. Move/rename it first if needed.
6. The new file is created in the **repository root**. Open in Excel, run Debug > Compile VBAProject, then the composite-key sample test.
7. Turn off the temporary **Trust access to the VBA project object model** permission afterwards.

If script execution is restricted by local/enterprise policy, follow the manual path below; avoid disabling those restrictions globally.

## Manual build (alternative)

1. Create a fresh workbook in Excel Desktop; Save As > Excel Macro-Enabled Workbook (.xlsm).
2. Alt+F11 > File > Import File. Import the **three** .bas modules listed above.
3. Run Debug > Compile VBAProject.
4. Alt+F8 > BuildUnifiedDashboard.
5. Save workbook. Remove only unused, empty template sheets. Do not delete Dashboard or Run History.

## Smoke test

Choose `samples/composite_before.csv` and `samples/composite_after.csv`, keys `CompanyID,InvoiceNo`, mode `EXACT`. Run comparison. Expected: 0 only A; 1 only B; 2 changed fields; 1 unchanged matched row; 0 issues. Generated report includes KPI Dashboard, Summary, Differences, Data Issues. Verify XLSX/PDF exports. After first run, workbook should contain Dashboard and Run History. See `tests/V06_ACCEPTANCE.md`.

## Status / safety

Earlier incremental XLSM versions were confirmed working by owner on Windows 11 + Excel 2019 / VirtualBox. **This clean build script has not yet been executed or validated there.** No precompiled XLSM has been released here. Do not publish populated Run History or personal Windows file paths with your portfolio file. Office 2019 is out of Microsoft support; use supported Office for customer production environments.

Code signing and final release packaging will follow successful runtime tests.
