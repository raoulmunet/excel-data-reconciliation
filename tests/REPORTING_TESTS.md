# v0.5 Reporting and Run History — manual acceptance tests

**Target:** Windows 11 / Office Excel 2019 Desktop, VirtualBox. **Status: not tested in Excel yet.** Earlier v0.4 is user-confirmed.

## Upgrade from the working v0.4 workbook

1. **Save a backup** of the working `.xlsm`.
2. In `Alt+F11`, remove ONLY the old `modUnifiedDashboard` module (answer **No** to Export only if you already saved a backup).
3. Download and import the **latest** `src/modUnifiedDashboard.bas`. Do not keep two versions with the same module name.
4. Import the **new** `src/modReportPresentation.bas`. Keep `modAdvancedReconciliation`, `modReconciliation`, and `modDashboard`.
5. Select **Debug > Compile VBAProject**. If compilation fails, record exact code line/error and stop.
6. Run `BuildUnifiedDashboard` from `Alt+F8`. Existing source paths, key columns, and mode are retained.
7. Use the provided composite CSVs with keys `CompanyID,InvoiceNo`, mode `EXACT`. Run comparison.

## Expected results

- A newly generated workbook has **four sheets**: `KPI Dashboard`, `Summary`, `Differences`, `Data Issues`.
- Summary stays consistent with v0.4 expected counts: A-only **0**, B-only **1**, field differences **2**, unchanged matched keys **1**, data issues **0**.
- KPI Dashboard: matched keys **3**, unchanged **1**, changed **2**, match rate **33.3%**, field differences **2**, data issues **0**, only in A **0**, only in B **1**. NOTE: in this fixture there are 3 common keys.
- `Differences` marks changed fields amber and B-only rows pale green. Page setup uses landscape on detail tabs, portrait on KPI and Summary.
- The tool workbook receives a new **Run History** sheet with status `SUCCESS`, timestamp, input paths, keys, mode, elapsed seconds.
- **Export XLSX** retains all four sheets; **Export PDF** includes the four sheets. Verify PDF readability and pagination.
- Run an invalid key name next; new generated report should not appear, `Run History` gets `FAILED / NO REPORT`, and the previous valid report should still be exportable.
- Cancelling an export should not create a file or change the report.
- Close report and try export: explanatory warning rather than crash.

## Calculation note

Match rate is **unchanged matched keys / all matched keys**. It does not count records found only in one source, duplicate keys or empty-key rows. `Changed keys` is a unique-key count, while `Field differences` can be higher because one row may contain multiple differences.

## Limits

- New code is not compiled/tested on this system yet; do not mark passing cases without recording results.
- `Run History` is stored in the macro workbook and persists only after the user **saves the .xlsm**. No personal client data should be committed to GitHub.
- The legacy advanced engine catches errors and displays message boxes; failures may be reported as `FAILED / NO REPORT` rather than including a detailed exception.
- If the report generation succeeds but styling fails, status becomes `SUCCESS / STYLE WARNING`, and the original report stays available.
- PDF print scaling needs visual review for wide and very large reports; this is not a performance-certified release.
