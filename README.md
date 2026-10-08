# Excel Data Reconciliation — VBA Business Tool

An Excel Desktop application for comparing two **UTF-8 CSV** datasets, finding missing/changed records and duplicate keys, generating KPI dashboards, and exporting reconciliation reports to XLSX and PDF. Designed as a practical Office/VBA portfolio project.

> **Release status:** Clean v0.6 **source and build script** available. Previous incremental builds were owner-tested on **Windows 11 + Excel 2019 in VirtualBox**. **The clean-build workflow and resulting XLSM are confirmed working by the project owner (Windows 11 / Excel 2019 / VirtualBox).** Individual edge cases, multi-version validation, and distribution QA remain pending. Do not advertise a finished, certified XLSM release until validation is complete.

## Main features

- Import two comma-delimited UTF-8 CSV datasets as **text**, preserving leading zeros and long identifiers.
- Compare records using one or more key columns (e.g. `CompanyID,InvoiceNo`).
- Configurable comparisons: `EXACT`, `TRIM`, `IGNORE_CASE`.
- Identify missing keys, changed fields, duplicate and missing keys.
- Generate a separate Excel report with `KPI Dashboard`, `Summary`, `Differences`, and `Data Issues`.
- Save the resulting report to XLSX or PDF.
- Keep `Run History` (input paths, run status, duration) in the macro workbook; save your workbook to persist history.

## Environment and compatibility

| Environment | Status | Notes |
|---|---|---|
| **Windows 11 + Excel 2019 Desktop in VirtualBox** | **Reference environment** | Earlier versions user-confirmed; clean v0.6 rebuild pending verification |
| Windows 11 + Excel 2021/2024/Microsoft 365 Desktop | Designed for, **not tested** | Must allow VBA macros |
| Windows + Excel 2016 Desktop | Compatibility target, **not tested** | No dynamic-array formulas required |
| Windows 10 + Desktop Excel | Potentially functional, **not tested** | Windows 10 mainstream support ended Oct 2025 |
| Excel Desktop 32-bit / 64-bit | Designed for both, **bitness not recorded** | No 32/64-bit Windows API declarations required |
| macOS Office | **Not supported/tested** | Windows ADODB.Stream / Scripting.Dictionary dependencies |
| Linux / LibreOffice / Wine | **Not supported** | Excel/VBA COM compatibility not guaranteed |
| Excel in browser / Office Online | **Not supported** | Browser Excel cannot run VBA |

**Other dependencies:** Excel Desktop with VBA. The current implementation uses Windows COM components `ADODB.Stream` and `Scripting.Dictionary` through late binding (no VBA reference checkbox required). No database, Outlook, Access, Word, internet access or external add-in is required at runtime.

**Support lifecycle:** Office 2016 and Office 2019 reached Microsoft support end on **14 October 2025**; they can still run locally but no longer receive regular fixes. Prefer supported Office for production client environments. Compatibility claims describe software behavior, not Microsoft support.

**Macro security:** Only run code you reviewed/trust. Do not globally disable macro protections. For automated compilation/build through PowerShell only, Excel requires temporary **Trust access to the VBA project object model**. Disable that permission again once the build succeeds. Organization policies may block this capability.

## Repository structure

```text
src/
  modAdvancedReconciliation.bas  # UTF-8 CSV parser and comparison engine
  modReportPresentation.bas      # KPI, formatting, run history
  modUnifiedDashboard.bas        # Dashboard UI and Excel/PDF export
build/
  Create-CleanWorkbook.ps1       # Excel Desktop COM builder
samples/
  composite_before.csv
  composite_after.csv
tests/
  V06_ACCEPTANCE.md
  REPORTING_TESTS.md
docs/
  BUILD_AND_RELEASE.md
```

Only the **three** `.bas` files in `src/` are necessary for the current application. Previous modules are preserved in [archived pre-clean source](https://github.com/raoulmunet/excel-data-reconciliation/tree/archive/pre-clean-build-2026-10-08).

## Build a clean XLSM — recommended

On **Windows 11 with Excel 2019 installed**:

1. Download the full repository with **Code > Download ZIP** and extract it into a normal local folder, e.g. `C:\Users\<user>\Documents\ExcelDataReconciliation\`.
2. Read `build/Create-CleanWorkbook.ps1`. It builds a brand-new workbook and **never overwrites an existing target file**.
3. In Excel, navigate to **File > Options > Trust Center > Trust Center Settings > Macro Settings**, and temporarily tick **Trust access to the VBA project object model**. This is only required for automatic module import.
4. Close Excel to release any existing COM processes.
5. Open **Windows PowerShell** in the repository folder and execute:

   ```powershell
   powershell.exe -NoProfile -File .\build\Create-CleanWorkbook.ps1
   ```

   If your PowerShell policy blocks local scripts, consult your organization's policy or follow the **manual** method below; do not disable security controls globally.
6. The script creates `ExcelDataReconciliation.xlsm` at the repository root with a fresh **Dashboard** worksheet and no legacy `Control Panel`. It imports precisely three VBA modules.
7. In Excel, open the new workbook and run **Debug > Compile VBAProject** (Alt+F11). Then test the composite CSV samples as described below.
8. Return to the Trust Center and untick **Trust access to the VBA project object model** after the build.

A macro-enabled workbook that came from the internet may be blocked by Excel depending on enterprise policy and Windows Mark-of-the-Web settings. Use your organization's approved trusted-file procedure rather than weakening protections.

### Manual method (no VBA-project automation permission)

1. Open Excel Desktop, create a **new** blank workbook, and save it as `ExcelDataReconciliation.xlsm`.
2. Press `Alt+F11`, choose **File > Import File**, and import exactly the three files in `src/` listed above.
3. Use **Debug > Compile VBAProject**. Resolve compile errors before running.
4. Press `Alt+F8` and run **`BuildUnifiedDashboard`**.
5. Remove only the unused blank default sheet after the new Dashboard is created. Do **not** remove `Dashboard` or `Run History`.
6. Save, close and reopen your workbook; test with synthetic CSV fixtures.

## Smoke test

Select:
- **Source A:** `samples/composite_before.csv`
- **Source B:** `samples/composite_after.csv`
- **Key columns:** `CompanyID,InvoiceNo`
- **Comparison mode:** `EXACT`

Expected summary: **Only A 0; Only B 1; Changed fields 2; Unchanged matched keys 1; Data issues 0.**

Expected KPI: **3 matched keys; 2 changed keys; match rate 33.3%.**

The macro workbook should contain **Dashboard** and, after one comparison, **Run History**. The generated report workbook should contain four tabs: **KPI Dashboard**, **Summary**, **Differences**, **Data Issues**. Confirm both XLSX and PDF exports manually. See [v0.6 tests](tests/V06_ACCEPTANCE.md) and [reporting tests](tests/REPORTING_TESTS.md).

## Limits and known engineering decisions

- CSV input is UTF-8 with comma separators. Legacy ANSI, semicolon-delimited files, or locale-specific CSVs are not guaranteed.
- Comparisons operate on text values. Configurable numeric/date tolerances and fuzzy matching are future features.
- Duplicate keys are excluded from unambiguous comparisons; file and memory size affect performance.
- The advanced engine handles its errors with Excel message boxes. A failure can be logged as `FAILED / NO REPORT` instead of a detailed exception.
- Worksheet reporting and PDF pagination may require refinement for large datasets.
- Never commit client data, local file paths, or populated `Run History` into the public repository.
- The clean workbook build has passed an owner-reported end-to-end smoke test in Excel 2019. Individual acceptance assertions, exported file review, and security/distribution checks remain to be documented.

## Roadmap

1. **v0.6:** Clean three-module architecture and reproducible workbook build — **basic clean-build and functional test confirmed by owner; release QA pending**.
2. **v0.7:** Automated acceptance checks, improved error reporting, stable sample release package.
3. **v1.0:** Validated XLSM release, screenshots, Upwork portfolio assets and versioned documentation.

## License

MIT License. All sample records are fictional.
