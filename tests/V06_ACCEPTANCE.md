# v0.6 upgrade and acceptance tests

**Reference environment:** Windows 11 in VirtualBox, Microsoft Excel 2019 Desktop.
**Status:** source published; v0.6 not yet verified inside Excel. v0.5 is backed up in Git branch `backup/v0.5-working`.

## Safe upgrade from working v0.5

1. Save and back up the existing **.xlsm** file before making changes.
2. Open Excel VBA editor (**Alt+F11**), right-click the existing `modUnifiedDashboard` and remove it. Keep an exported backup of that module.
3. Import the updated [`src/modUnifiedDashboard.bas`](../src/modUnifiedDashboard.bas) through **File > Import File**. Keep `modAdvancedReconciliation` and `modReportPresentation` imported. The legacy `modDashboard` / `modReconciliation` modules may remain.
4. Run **Debug > Compile VBAProject**.
5. Press **Alt+F8** and run `BuildUnifiedDashboard`. The existing worksheet `Reconciliation v0.4` should automatically be **renamed** to `Dashboard` while retaining values in B6, B8, B10, B12.
6. If the old worksheet was already deleted, `BuildUnifiedDashboard` creates a fresh `Dashboard`. Re-enter paths, key columns and mode.
7. Run the test below and verify outcomes.

**Do not delete `Run History`.** The old `Control Panel` sheet is no longer needed for the unified workflow; it is safe to omit it from a new distributable workbook. Do not delete or rename the underlying legacy modules until other macros have been reviewed.

## Functional acceptance tests

| Test | Expected result | Verification |
|---|---|---|
| Compile VBA project | No compilation errors | Pending |
| Migrate previous worksheet | `Reconciliation v0.4` becomes `Dashboard` | Pending |
| Preserve settings | B6/B8 paths, B10 key, B12 mode retained | Pending |
| Run after deleting old Control Panel | Works without errors | Pending |
| Run composite fixture in EXACT mode | 0 A-only, 1 B-only, 2 field differences, 1 unchanged matched key, 0 data issues | Pending |
| KPI Dashboard | 3 matched keys, 2 changed keys, 1 unchanged, 33.3% match rate | Pending |
| Run History | New successful row and duration | Pending |
| XLSX export | Four sheets; no macro code inside report | Pending |
| PDF export | Readable KPI, Summary, Differences and Data Issues | Pending |
| Rebuild again | One Dashboard, unchanged configuration | Pending |
| Fresh workbook without legacy Control Panel | Unified workflow runs | Pending |
| Open/close/report errors | Messages useful; no stale report exported | Pending |

## Packaging

For a clean new release workbook, import:
- `src/modAdvancedReconciliation.bas`
- `src/modReportPresentation.bas`
- `src/modUnifiedDashboard.bas`

The v0.6 unified interface calls **only** the advanced and presentation modules. Older `modReconciliation.bas` and `modDashboard.bas` are maintained in source control for historical support but are not required by the new interface.

After validation, create and save a clean `ExcelDataReconciliation.xlsm` in Windows Excel 2019, then remove any demo history or personal file paths before distribution. Never bypass the user's macro security settings. A signed workbook requires a separate signing process; this repository does not include a certificate.

## Recovery

If v0.6 shows an error, use your saved v0.5 `.xlsm` copy or retrieve the previous VBA modules from the GitHub `backup/v0.5-working` branch. Report the error text and the VBA line highlighted by Debug.
