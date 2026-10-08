# Advanced reconciliation acceptance tests (v0.3 preview)

**Status (2026-10-08):** project owner confirmed a successful basic run of the advanced macro after selecting the correct composite-key CSV fixtures on Windows 11 + Office 2019 (VirtualBox). Previously observed `Unknown key column: CompanyID` was caused by selecting the older `CustomerID` fixtures. Exact expected metrics, comparison modes and edge cases have **not** been separately verified.

## Installation
1. Open the existing `.xlsm` workbook in Excel 2019 Desktop.
2. Import `src/modAdvancedReconciliation.bas` alongside the existing modules through the VBA editor **File > Import File**.
3. Use **Debug > Compile VBAProject**. Note and report errors before enabling production use.
4. Launch `RunAdvancedReconciliation` with **Alt+F8**. Select `composite_before.csv` and `composite_after.csv`.
5. At the key prompt, enter `CompanyID,InvoiceNo`.

## Expected results

| Mode | Only A | Only B | Changed fields | Unchanged matched rows | Issues |
|---|---:|---:|---:|---:|---:|
| EXACT | 0 | 1 | 2 | 1 | 0 |
| TRIM | 0 | 1 | 2 | 1 | 0 |
| IGNORE_CASE | 0 | 1 | 1 | 2 | 0 |

For EXACT and TRIM, differences: `INV-0001/Comment` (`Alpha` vs ` alpha `) and `INV-0002/Amount` (`55.50` vs `58.50`). In IGNORE_CASE, only the amount changes.

## Edge cases to test

- Key `001` must stay distinct from `1`.
- Quoted commas, escaped quotes, and multiline quoted CSV fields.
- BOM encoded UTF-8 and accented characters.
- Duplicate composite keys on either side.
- Empty key component.
- Header columns reordered in B (comparison should be header based).
- Wrong number of CSV fields should display a parsing error.
- Missing source file or invalid comparison mode should produce an explanatory error.
- Excel 32-bit and 64-bit independently (bitness not yet confirmed).
- Very large datasets: no performance benchmarks yet.

**Important:** Test using synthetic data. Do not use private client spreadsheets for initial verification.

## Troubleshooting a missing key column

An error such as `Unknown key column: CompanyID` indicates that Source A's parsed header does not contain that column. The original customer fixtures contain `CustomerID` instead. Confirm you selected the **composite** fixture files and typed `CompanyID,InvoiceNo` exactly. The 2026-10-08 source update includes a more informative error listing Source A and its detected columns, plus BOM normalization. Re-import the latest module before retesting. This fix is not yet runtime verified.

## Execution log

- **2026-10-08 — Basic advanced-flow smoke test: PASS (owner-reported).** Correct composite CSV fixtures selected; macro ran successfully. Earlier wrong-file selection explained the missing key. Specific report counts and individual comparison modes not confirmed; the remaining cases above are pending.
