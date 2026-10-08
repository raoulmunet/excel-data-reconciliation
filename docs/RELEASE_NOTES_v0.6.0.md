# Excel Data Reconciliation — v0.6.0 (Portfolio Demo)

**Release type:** portfolio demonstration / Windows Excel Desktop VBA utility.

The downloadable release package should be named:
`ExcelDataReconciliation_v0.6_Portfolio_Package.zip`

## Highlights

- Unified Excel dashboard with two CSV inputs and configurable key columns.
- Composite-key reconciliation using a UTF-8 text-preserving CSV parser.
- Three comparison modes: EXACT, TRIM, IGNORE_CASE.
- Report workbook with KPI Dashboard, Summary, Differences, Data Issues.
- XLSX and PDF report export.
- Run History in the macro-enabled workbook.
- Clean application architecture comprising three VBA modules.

## Tested environment

The project owner confirmed successful basic operation of a freshly generated XLSM workbook on **Windows 11 (VirtualBox) with Excel 2019 Desktop** on **2026-10-08**. Specific 32/64-bit Office architecture, individual comparison modes, and comprehensive edge cases have not been independently verified. Excel 2019 is out of Microsoft support.

## Package contents

- `ExcelDataReconciliation.xlsm`
- `samples/composite_before.csv`
- `samples/composite_after.csv`
- `START_HERE.md`

## Quick start

1. Download and unzip the package locally.
2. Read `START_HERE.md`, review VBA code and comply with your macro security policy.
3. Open `ExcelDataReconciliation.xlsm` in **Excel Desktop for Windows**.
4. Select the two bundled composite CSV files in Dashboard.
5. Enter `CompanyID,InvoiceNo` for key columns, select `EXACT`, and run.
6. Inspect KPI Dashboard / Summary / Differences / Data Issues. Optionally export XLSX or PDF.

Expected demo summary: **Only A: 0; Only B: 1; Changed fields: 2; Unchanged matched records: 1; Data issues: 0**.

## Compatibility and caveats

Excel for the web cannot run VBA. Mac, Linux, LibreOffice, and other Office Desktop versions are not validated. Source paths and Run History may expose local folder names if populated; this demo package was inspected for source paths in worksheet cells, but recipients should still inspect document properties and embedded VBA as part of release security review. Avoid uploading any client/private datasets.

## Source & documentation

Repository: https://github.com/raoulmunet/excel-data-reconciliation

See `README.md` for requirements, `docs/BUILD_AND_RELEASE.md` for source builds, and `tests/V06_ACCEPTANCE.md` for the validation checklist.

**Release publishing status:** These notes are prepared for GitHub Releases. The ZIP must be uploaded as a binary asset through the GitHub Releases interface before the release is considered published.
