# Test cases (manual Excel execution pending)

## Fixture A vs B
Select `samples/customers_before.csv` and `samples/customers_after.csv`, key `CustomerID`.

Expected **Summary**:
- Only in A: **1** (`103`)
- Only in B: **1** (`104`)
- Field differences: **2** (`100/Balance: 50 -> 55`, `101/City: Paris -> Lyon`)
- Matching records unchanged: **1** (`102`)
- Data issues: **1** (duplicate key `105` in B, excluded)

## Additional cases to execute
1. Swap source column order: results identical.
2. Duplicate key in A: one issue; duplicate key excluded from matching.
3. Blank nonempty row key: one `<EMPTY KEY>` issue.
4. Extra header in B: explicit error.
5. Nonmatching key column: explicit error.
6. Cancel the file picker: no output workbook.
7. Same source path selected twice: explicit error.
8. Empty file and header-only file: clear error.
9. CSV with quoted comma / quotes / accented UTF-8 text: inspect native Excel import behavior.
10. Leading-zero key `0012` vs `12`: known potential coercion; not a passing criterion until text-preserving parser.
11. Excel 32-bit and Excel 64-bit: import/compile/run separately.

All tests are **pending**. Do not report them as passed until results are recorded from Excel.

## Dashboard UI v0.2 test plan (not executed)

1. Import both `modReconciliation.bas` and `modDashboard.bas`; run **Debug > Compile VBAProject**; expect no compile errors.
2. Execute `BuildDashboard`; expect a **Control Panel** worksheet with Browse A, Browse B, Run reconciliation, Save report buttons.
3. Select sample CSVs; expect paths in `B6` and `B8`; enter `CustomerID` in `B10`.
4. Click Run reconciliation; expect a separate report workbook, plus a completion message in `B14`.
5. Verify Summary metrics against the fixture counts above (1 A-only, 1 B-only, 2 changed fields, 1 unchanged matched record, 1 issue).
6. Return to Control Panel, click Save report; choose a temporary `.xlsx` file; reopen it and check three expected sheets.
7. Click Browse and cancel: existing paths should remain unchanged.
8. Enter a nonexistent path: visible validation error, no new report.
9. Run `BuildDashboard` again: buttons regenerated; `B6`, `B8`, `B10` settings preserved.
10. Close the report workbook and attempt Save report: expect explanatory error, not a crash.
11. Run the comparison twice; Save report should target the second result.
12. Disable macro execution: document security restrictions without bypassing policy.

**Execution evidence:** awaiting tests on Windows 11 + Office 2019 VirtualBox. Do not label these tests as passed before the owner confirms them.
