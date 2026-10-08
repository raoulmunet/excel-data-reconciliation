# Architecture and engineering decisions

## Workflow
1. Select two source CSV files and a unique key column.
2. Open them as temporary Excel workbooks (read-only intent; never save source changes).
3. Validate equivalent header sets, unique nonempty header names, and key column.
4. Load each used range into a VBA array; create late-bound dictionaries mapping normalized key -> row.
5. Mark duplicate and empty keys as data issues and exclude ambiguous keys from matching.
6. Compare fields by header name, not source column position.
7. Output Summary, Differences, and Data Issues sheets to a new workbook.
8. Close source workbooks and restore Excel application settings even after an error.

## Decisions
- **Arrays and dictionaries** instead of per-cell worksheet lookups for reconciliation.
- **Late binding** avoids an explicit Microsoft Scripting Runtime reference.
- **Report separation** keeps validation issues distinct from business differences.
- **Fail-fast header checks** avoid misleading comparisons.
- **Non-destructive** behavior: sources never saved.
- **Native Excel CSV import** was chosen for the MVP for simplicity but has a significant drawback: Excel may coerce keys, dates and decimal values. Plan a deterministic text parser in v0.2.

## Known risks
- UTF-8 BOM / locale-dependent CSV parsing.
- Leading zeros and numeric precision lost during Excel import.
- Duplicate counts currently count one issue per duplicated key, not per duplicate row.
- No normalized numeric comparisons, configurable delimiters or tolerance rules.
- Sheet row ceiling (1,048,576) and in-memory arrays constrain file sizes.
- No Excel runtime tests have been run yet.
