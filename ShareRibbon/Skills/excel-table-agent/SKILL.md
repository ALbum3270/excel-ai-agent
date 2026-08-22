---
name: excel-table-agent
description: Use for Excel tasks that need table understanding, cleanup, calculation, formula repair, transpose, pivot tables, charts, reporting, or multi-step spreadsheet automation.
application: Excel
default_for_application: true
tags: excel, spreadsheet, table, formula, chart, pivot, clean, transpose, statistics, data-analysis
allowed-tools: ReadRange, ApplyFormula, WriteData, FormatRange, CreateChart, CreateSheet, DeleteSheet, RenameSheet, CopySheet, InsertRowCol, DeleteRowCol, HideRowCol, ProtectSheet, CleanData, SortData, FilterData, RemoveDuplicates, ConditionalFormat, MergeCells, AutoFit, FindReplace, CreatePivotTable, TransformData, DataAnalysis, GenerateReport, PythonCompute, DiscoverOfficeCapability, OfficeObjectOperation
intent_types: data_analysis, formula, chart, table_format, data_clean, transform
---

# Excel Table Agent

Use this skill when the user asks Excel to do real spreadsheet work from natural language, especially:

- organize or beautify a table
- calculate totals, averages, rankings, rates, commissions, KPIs, or custom formulas
- repair an incorrect formula or algorithm
- generate a chart from selected data
- transpose rows and columns
- clean duplicates, blanks, spaces, inconsistent text, or obvious data quality issues
- create a pivot table, summary table, or report sheet

## Operating Rules

1. First read the current Excel context: workbook, sheet, selection address, used range, headers, sample rows, data types, numeric columns, text columns, formula cells, blanks, duplicates, and candidate target area.
2. Prefer native JSON tools over VBA. Use `ExecuteVBA` only when registered tools cannot express the task.
3. Treat a selected multi-cell range with content as the primary data range. A single empty cell outside the detected table is only a possible write/output target; it must not replace an observed `TableRegion`/`UsedRange` as the source for analysis or charts.
4. Do not ask the user for data that the add-in can observe from Excel.
5. For write operations, produce a plan with target ranges and expected effects before acting when the task is medium or risky.
6. After each operation, observe the sheet state: changed range, formulas, chart count, row/column shape, or generated summary.
7. If execution fails, repair the tool parameters using the observation. Do not fall back to a long chat explanation unless execution is impossible.
8. When the user does not specify a calculation engine, use `PythonCompute` only when native tools cannot express the calculation. First use `ReadRange` for the smallest necessary range, pass its JSON data to Python, then write the approved result with native tools. Python must not access Excel, files, network, or child processes.
9. If the user explicitly requests `Python` or `PythonCompute`, you must not substitute `DataAnalysis`, formulas, pivots, or another engine. Use `ReadRange -> PythonCompute -> WriteData`; controlled `PythonCompute` needs no approval, and `CreateSheet` is used only when the destination sheet does not already exist.
10. If a follow-up only corrects the output sheet or target range, preserve the preceding task's source range, calculation method, grouping fields, and aggregation. Do not treat the active output sheet as the new source unless the user explicitly changes the source.
11. When a registered high-level tool directly covers the request, call that tool rather than wrapping its name or parameters inside `OfficeObjectOperation`. For a long-tail Excel object operation not covered by a high-level tool, call `DiscoverOfficeCapability` first in the same task. Copy the returned executable `MemberId` exactly into `OfficeObjectOperation`; never invent or shorten one.
12. A request to answer without writing is read-only, not context-free. When an exact answer depends on workbook values and the context contains only a preview, use `ReadRange` for the smallest complete necessary range and answer from the returned JSON. Do not call mutation tools, extrapolate from samples, estimate missing rows, or invent values.
12. Every mutating `OfficeObjectOperation` must include an observable `expectedEffects` object or batch `successCriteria`. A host call returning without an exception is not proof of success.

## Declarative Object Operations

- Use canonical refs rooted at `Excel:workbooks/active`, for example `Excel:workbooks/active/worksheets/Sales/ranges/D2%3AD25` or `Excel:workbooks/active/worksheets/Sales/chartobjects/Chart%201`.
- Encode `/`, `!`, spaces, and other reserved characters inside path segments. Do not put COM objects in plans, memory, or tool arguments; use `{ "ref": "Excel:..." }` when another Office object is required as an argument.
- Build batches with `schemaVersion=1.0`, `appType=Excel`, unique operation IDs, and actions limited to `get/set/invoke/create/delete/collection_item`.
- Treat `VERIFY_FAILED`, `OBSERVATION_FAILED`, and `PARTIAL_APPLY` as real failures. Do not repeat a mutating batch automatically when Excel may already have changed.
- Delete, clear, close, and overwrite operations require approval through the existing SafetyGate.

## Tool Preferences

- Table formatting: `FormatRange`, `AutoFit`, `ConditionalFormat`
- For `FormatRange`, always express the requested effects as explicit properties. Use `numberFormat` for numeric/date/percentage/currency display; do not infer those effects from a preset style name.
- Formula generation or repair: `ApplyFormula`
- Summary/statistics: `DataAnalysis`, `ApplyFormula`, `CreatePivotTable`
- Chart generation: `CreateChart` after selecting chart type from data shape
- Row/column transpose: `TransformData` with `operation=transpose`
- Cleaning: `CleanData`, `RemoveDuplicates`, `FindReplace`
- Report output: `GenerateReport`, `WriteData`, `CreateChart`
- New result sheet: `CreateSheet`, then `WriteData`; use `RenameSheet` only for an explicit rename request
- A worksheet name such as `Summary`, `汇总`, or `Report` does not imply report content. If the user only asks to add/create a named worksheet, use `CreateSheet` only and do not use `GenerateReport`.
- Complex calculation: `PythonCompute`, then inspect its JSON result and use `CreateSheet`/`WriteData` to write the verified result
- Structured input: `ReadRange` before `PythonCompute`; do not reconstruct a large table from the text preview

## Chart Selection Heuristics

- Trend over time: line chart
- Category comparison: column or bar chart
- Share of total with few categories: pie chart
- Relationship between two numeric columns: scatter chart
- Multiple numeric series by category: clustered column chart

## Formula Repair Heuristics

When asked to fix an algorithm or formula:

1. Inspect the source columns and the current formula if present.
2. Identify references, relative/absolute addressing, row offsets, and likely fill-down range.
3. Generate the corrected formula for the first data row.
4. Apply it to the target range with fill-down when appropriate.
5. Observe whether Excel reports formula errors such as `#VALUE!`, `#REF!`, `#N/A`, or empty results.
