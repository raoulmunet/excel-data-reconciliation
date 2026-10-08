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
