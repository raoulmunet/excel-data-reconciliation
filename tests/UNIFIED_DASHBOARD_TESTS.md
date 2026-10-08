# v0.4 Unified Dashboard — Windows Excel 2019 test plan

**Status: not yet executed in Microsoft Excel.** Tested v0.2 dashboard and basic v0.3 engine are preserved; this new interface needs validation on Windows 11 / Office 2019 (VirtualBox).

## Installation

1. Back up the working `.xlsm` first.
2. In VBA editor (`Alt+F11`), keep existing `modReconciliation`, `modDashboard`, and `modAdvancedReconciliation` modules.
3. Download and import **`src/modUnifiedDashboard.bas`**, using **File > Import File**.
4. Run **Debug > Compile VBAProject**. Stop and report the exact error/line if compilation fails.
5. In Excel press `Alt+F8` and run `BuildUnifiedDashboard` (once to create the sheet).
6. A new sheet **Reconciliation v0.4** will appear without replacing the older **Control Panel**.
7. Select **Source A** and **Source B**, specify `CompanyID,InvoiceNo` in the **Key columns** field, select `EXACT` from the dropdown.
8. Press **RUN COMPARISON**, review the generated Summary / Differences / Data Issues workbook.
9. Return to the new dashboard sheet to click **Export XLSX** or **Export PDF**.

## Acceptance criteria

| Case | Expected | Actual |
|---|---|---|
| Compile all imported modules | No syntax/compile error | Pending |
| `BuildUnifiedDashboard` | New sheet, 5 working buttons | Pending |
| Browse A/B; cancel picker | Path updated / cancel leaves path unchanged | Pending |
| Composite keys + EXACT + samples/composite_* | 0 A-only, 1 B-only, 2 field diffs, 1 unchanged, 0 issues | Pending |
| Mode TRIM | 0 A-only, 1 B-only, 2 diffs, 1 unchanged, 0 issues | Pending |
| Mode IGNORE_CASE | 0 A-only, 1 B-only, 1 diff, 2 unchanged, 0 issues | Pending |
| XLSX export | File saved with 3 sheets and correct summary | Pending |
| PDF export | Opens with all 3 report sheets present and legible | Pending |
| Missing or same-file sources | Clear warning, no new report | Pending |
| Run failure after success | Previous valid report remains available for export | Pending |
| Rerun after closing prior report | New output tracked correctly | Pending |
| Rebuild unified sheet | Paths, key columns, comparison mode remain | Pending |
| v0.2 `BuildDashboard` still works | Older control panel unchanged | Pending |

## Notes

- The dashboard uses worksheet shapes, not ActiveX controls.
- It calls `AdvancedReconcileCsvFiles` from **modAdvancedReconciliation**, rather than the older `ReconcileCsvFiles`.
- Its currently selected report is kept in VBA memory only. Closing/resetting Excel requires a new reconciliation before export.
- The PDF export uses Excel's built-in `ExportAsFixedFormat`, exporting the entire report workbook; layout may require print-setup refinements for large reports.
- CSV input is expected to be UTF-8 and comma-delimited. See the main README for limitations.
- The latest v0.4 module has no runtime execution evidence yet; GitHub publication is not an Excel test.

## Known limitations / next improvements

- No automatic configuration file or error log persistence.
- Report is generated in a separate workbook and export requires returning to the Control Panel.
- The advanced engine catches its own errors and uses dialogs; a future revision should return an explicit success flag/report handle.
- No configurable numeric/date tolerance or advanced formatting yet.
