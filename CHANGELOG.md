# Changelog

All notable changes to this project are documented in this file.

## [Unreleased]

### Added
- **ISSUE-CELLSGO-302**: Chart export functionality — `converter.ExportChartToSink`, `converter.ExportChartToBytes`, and `converter.ExportChartToFile` export charts from workbooks to various image formats. Supported formats: PNG (`ChartExportFormatPNG`), JPEG (`ChartExportFormatJPEG`), SVG (`ChartExportFormatSVG`), and PDF/EMF vector format (`ChartExportFormatPDF`). `ChartExportOptions` configures format, dimensions (width/height), and JPEG quality. The output shape (file, `io.Writer`, or in-memory bytes) is chosen by picking the sink, matching the existing `converter.Convert` pattern. New sentinel: `ErrUnsupportedFormat` for unknown format strings.
- **ISSUE-CELLSGO-301**: Data validation support — `editor.AddDataValidation(cellRange, actions...)` adds validation rules to cell ranges, `editor.InValidation(index, actions...)` modifies existing validations, `editor.DeleteValidation(index)` removes them. Validation types: `ValidationTypeWholeNumber`, `ValidationTypeDecimal`, `ValidationTypeList`, `ValidationTypeDate`, `ValidationTypeTime`, `ValidationTypeTextLength`, `ValidationTypeCustom`. Operators: `OperatorTypeBetween`, `OperatorTypeEqual`, `OperatorTypeNotEqual`, `OperatorTypeLessThan`, `OperatorTypeLessOrEqual`, `OperatorTypeGreaterThan`, `OperatorTypeGreaterOrEqual`. Actions: `WithValidationType`, `WithValidationOperator`, `WithValidationFormula1/2`, `WithValidationList`, `WithValidationInCellDropDown`, `WithValidationIgnoreBlank`, `WithValidationShowInput`, `WithValidationShowError`, `WithValidationAlertStyle`, `WithValidationErrorTitle/Message`, `WithValidationInputTitle/Message`. New sentinels: `ErrValidationNotFound`, `ErrInvalidValidationType`, `ErrInvalidOperatorType`.
- **ISSUE-CELLSGO-301**: Conditional formatting support — `editor.AddConditionalFormatting(sheetID, cellRange, actions...)` adds conditional formatting rules, `editor.InConditionalFormatting(index, actions...)` modifies existing rules, `editor.DeleteConditionalFormatting(index)` removes them. Rule types: color scales (`WithColorScale`), data bars (`WithDataBar`), icon sets (`WithIconSet`), cell value comparisons (`WithCellValueRule`), formula expressions (`WithExpressionRule`), above/below average (`WithAboveAverageRule`), top/bottom N (`WithTop10Rule`). Icon sets: `IconSetTrafficLights31`, `IconSetArrows3`, `IconSetRating5`, and more. New sentinels: `ErrConditionNotFound`, `ErrInvalidFormatConditionType`, `ErrInvalidIconSetType`.
- **ISSUE-CELLSGO-301**: Query support for data validation and conditional formatting — `query.ValidationCount`, `query.ValidationInfoAt`, `query.AllValidations` read validation metadata; `query.ConditionalFormattingCount`, `query.ConditionalFormattingInfoAt`, `query.AllConditionalFormattings` read conditional formatting metadata. Both return Go-native structs (`ValidationInfo`, `ConditionalFormattingInfo`, `ConditionInfo`) without exposing engine objects.
- **ISSUE-CELLSGO-301**: `examples/validation-cf/` — comprehensive example demonstrating data validation (whole number, list, decimal validations with error messages and input prompts) and conditional formatting (color scales, data bars, icon sets, cell value rules, expression rules, top 10 rules). Shows query API usage for reading back validation and conditional formatting metadata.
- **ISSUE-CELLSGO-301**: Internal helpers in `internal/aspose/cells/validation.go` and `internal/aspose/cells/conditional_format.go` for validation type/operator/alert resolution, conditional formatting condition type/icon set type resolution, and engine object access.

### Fixed
- **ISSUE-CELLSGO-279**: Referencing a worksheet by a name that does not exist now returns `ErrWorksheetNotFound` instead of crashing the process. The binding's `Get_String(name)` returns `err=nil` with a dangling native handle for missing names; every by-name lookup (`editor.WithRenameWorksheet`, `editor.InWorksheet`, `transfer.ExportRangeToJson`, `transfer.ExportWorksheetToJson`) now iterates the collection via the new `internal/aspose/cells.WorksheetByName` helper.
- **ISSUE-CELLSGO-279**: Committed the `examples/` and GitHub Actions CI that v26.8.0 documented but never shipped (both were hidden by `.gitignore`). The examples were rewritten against the released API, and `.gitattributes` (`* text eol=lf`) was added so `gofmt` and tests behave identically on Linux and Windows.
- **ISSUE-CELLSGO-279**: `manipulator.Merge` no longer silently succeeds with zero input sources (which produced a meaningless empty workbook); it now returns `ErrNoSources`. Errors raised while reading or combining a source now carry the source index (`source 0: …`), and `manipulator.Split` errors carry the offending sheet name (`sheet "…": …`), so failures in loops are traceable.
- **ISSUE-CELLSGO-279**: Exporting an empty worksheet / cell range with `ExportWorksheetToJson` / `ExportRangeToJson` previously returned a successful 0-byte result; it now writes a valid empty JSON array (`[]`) so callers always receive a parseable document.
- **ISSUE-CELLSGO-279**: `datasource.FilePathSink` now creates missing parent directories (matching `FolderSink`), so writing to a not-yet-existing output folder no longer fails.
- **ISSUE-CELLSGO-279**: Fixed flaky behavior tests caused by evaluation-mode load-time worksheet-name corruption. The engine rewrites a random worksheet's name to garbage bytes on ~2% of `NewWorkbook_Stream` loads (any sheet; not present in the saved bytes, so the same file can load clean once and corrupt later). Name-sensitive tests (`TestMergeCombinesWorkbooks`, `TestSplitProducesOneOutputPerSheet`, `TestExportEmptySheetWritesEmptyJson`, `TestImportCSVLandsInCells`) now re-run on freshly generated input via a `retryStable` helper until the engine yields clean names; a genuine deterministic regression fails every attempt and still fails the test. Documentation (`WithSheet`, README, `docs/design.md`) now describes the mechanism accurately.
- **ISSUE-CELLSGO-294**: Fixed doc comments and examples in `editor` (`engine.go`, `types.go`) and `docs/editor.md` that referenced non-existent `InSheet` / `SetCellStyle` (the actual API is `InWorksheet` / `SetStyle`); the `EditSpreadsheet` doc example passed an `int` to the `string`-typed `WithActiveSheet` and would not compile. Examples now use real identifiers.
- **ISSUE-CELLSGO-294**: README Quick Start now shows and explains the `_ "…/register"` blank import (required only when an entry point resolves the format from a file extension / `formats.Get`), and a new **Deprecation schedule** section documents that the 18 legacy wrappers are removed in v27.0.0, with a per-package list.
- **ISSUE-CELLSGO-294**: `editor.SetCellValue` / `editor.SetValue` now apply a date number format when the value is a `time.Time`. Previously the engine's `NewObject_Date` path wrote only a numeric serial without a date format, so a written date was displayed and read back as a bare number (e.g. `45293`); now the cell saves and reloads as a real date (`2024-01-02`), which the `query` package recognizes as `KindDateTime`.
- **ISSUE-CELLSGO-297**: `core.SetLicense` now returns an `error` and wraps `errors.ErrLicenseInvalid` when the license cannot be created or applied, instead of silently ignoring the failure. Callers using the statement form (`core.SetLicense(path)`) keep compiling unchanged.
- **ISSUE-CELLSGO-297**: `editor.InWorksheet` with an out-of-range integer sheet index now returns `ErrInvalidSheetID` instead of reaching the engine with a bad index.

### Added
- **ISSUE-CELLSGO-294**: `query` package — the read counterpart of `transfer`. It loads a spreadsheet from a `datasource.DataSource` and returns Go-native data: typed `CellValue` grids (`ReadCell`, `ReadRange`, `ReadWorksheet`), merged regions (`ReadMergedCells`), sheet names (`SheetNames`), and used-range dimensions (`Dimensions`). `CellValue` pairs a value with a `CellKind` (empty/text/int/float/bool/date-time/error) and exposes matching accessors; options are `WithSheet`, `WithSheetIndex` (index-based, immune to evaluation-mode name corruption; the default targets the first sheet by index), and `WithTrimSpace`. `CellRef` / `Area` plus `ParseCellRef` / `ParseArea` provide Excel-style cell addressing.
- **ISSUE-CELLSGO-294**: `editor.EditSpreadsheetToSink(source, sink, actions...)` — sink-based form of `EditSpreadsheet`, letting the caller choose the output shape (file, writer, or bytes); `EditSpreadsheet` is now a thin wrapper over it.
- **ISSUE-CELLSGO-294**: `editor.SetFormula(row, col, formula)` worksheet action and `editor.CalculateAll()` workbook action for writing and recalculating formulas.
- `errors.ErrInvalidCellRef` sentinel for unparsable cell references (e.g. "B3").
- `errors.ErrInvalidRange` sentinel for malformed or reversed cell ranges.
- `errors.ErrWorksheetNotFound` sentinel for classifying missing-worksheet failures with `errors.Is`.
- `errors.ErrDataSourceNil` and `errors.ErrDataSinkNil` sentinels for classifying nil source/sink failures with `errors.Is`.
- `errors.ErrNoSources` sentinel for classifying empty-merge failures with `errors.Is`.
- `datasource.NewWriterSink(w io.Writer)` and `datasource.NewZipSink(zw *zip.Writer)` constructors for the writer/zip sinks.
- **ISSUE-CELLSGO-279**: Composite entry points — `converter.Convert`, `manipulator.Merge` / `manipulator.Split`, and sink-based `transfer` exports (`ExportWorksheetToJson`, `ExportRangeToJson`, `ExportSpreadsheetToXml`) and imports (`ImportCSV`, `ImportJsonData`, `ImportXMLData`) with functional Options (`WithSheet`, `WithStartCell`, `WithEndCell`, `WithXMLMap`, `WithBeginCell`, `WithConvertNumeric`, `WithSeparator`). One composite per capability: the output shape (file, `io.Writer`, bytes, folder, or zip archive) is chosen by picking a `datasource.DataSink`.
- **ISSUE-CELLSGO-297**: `internal/aspose/cells.LoadStable(source, attempts, verify)` — retry a load until the verification callback passes, discarding engine-corrupted loads (evaluation-mode load-time name/value corruption) and reporting a wrapped error when every attempt fails. Tests demonstrate the recipe: read → verify → reload on failure.
- **ISSUE-CELLSGO-297**: `transfer.WithSheetIndex(i int)` — index-based sheet targeting for the transfer imports/exports, immune to evaluation-mode name corruption; the transfer default sheet is now the first worksheet **by index** (0), aligned with `query`. `errors.ErrLicenseInvalid` sentinel classifies license failures.
- **ISSUE-CELLSGO-298**: Structured table read/write — `query.ReadRows[T]` maps a worksheet's used range into `[]struct` and `editor.WriteRows[T]` writes `[]struct` back, both sharing one column-mapping rule (`excel:"name"` tag, else field name, `excel:"-"` to skip; header match is case-insensitive). `ReadRows` supports string / int / uint / float / bool / `time.Time` fields, decodes an empty cell to the zero value, and reports a missing column as the new `errors.ErrColumnNotFound`; `WriteRows` offers `WithWriteHeader`, `WithSheetIndex`, and `WithSheet` (default first sheet by index). The reflection mapping lives in the shared `internal/rows` package so the two sides cannot drift.
- **ISSUE-CELLSGO-298**: `editor.toObject` now also converts `uint` / `uint8` / `uint32`, and `query.CellKind` gains a `String()` method for readable error messages.
- **ISSUE-CELLSGO-299**: Named ranges, cell comments, and workbook encryption. `query.NamedRanges` lists a workbook's defined names as `NamedRange{Name, RefersTo, Area}` (area resolved best-effort), `query.ReadNamedRange` reads a name's cells as a grid, and `query.ReadCellComment` returns a cell's comment note (empty when none). The editor counterparts are `DefineNamedRange` (idempotent; returns `ErrInvalidRange` on reversed coordinates), `SetCellComment`, and `ClearComments`. `editor.Encrypt` encrypts the saved workbook (strong AES, `Workbook.SetEncryptionOptions`) so it requires the password to open; loading encrypted files back needs a password at load time, which the toolkit loader does not yet expose. Comment reads iterate the engine's comment collection because the binding's `Cell.GetComment` returns a dangling handle (hard crash on use) for a comment-free cell. New sentinel `errors.ErrNameNotFound`; `internal/aspose/cells.CellRef` gains `AbsoluteString()` for `$A$1`-style references.
- **ISSUE-CELLSGO-300**: Chart support in the `editor` DSL — add, delete, and modify charts without touching the engine binding. `editor.AddChart(chartType, dataRange, byColumn, topRow, leftColumn, bottomRow, rightColumn, actions...)` creates a chart (returning it for the chained actions), `editor.InChart(index, actions...)` targets an existing one, and `editor.DeleteChart` / `editor.DeleteAllCharts` remove them. The rectangle is a required parameter rather than an optional action and there is no default placement, so a chart is never drawn somewhere the caller did not ask for; `WithChartBounds` remains for moving a chart that already exists. `ChartAction`s cover the title (`WithChartTitle`, `HideChartTitle`), the built-in style (`WithChartStyle`, 1..48), the type (`WithChartType`), the legend (`WithChartLegend`, `WithChartLegendPosition`), placement (`WithChartBounds`), and the data (`WithChartDataRange` to re-point wholesale, plus `WithChartSeries` / `RemoveChartSeries` / `ClearChartSeries` / `WithChartCategoryData` for series-level work). Chart types and legend positions are toolkit-native `string` enums — 34 curated `ChartType` constants plus raw-name fallback to reach all 81 engine types, matched case- and punctuation-insensitively — so `asposecells.ChartType` is never exposed. Validation is deliberately strict because the engine validates nothing: an out-of-grid, reversed, single-cell, or single-row range yields a silently empty chart rather than an error, and a style number above 48 is silently clamped to 48, so ranges are parsed and grid-bounds-checked and styles range-checked before the engine sees them (「能报错就报错,不静默降级」). New sentinels `errors.ErrChartNotFound`, `errors.ErrInvalidChartType`, `errors.ErrInvalidChartStyle`, `errors.ErrInvalidChartPosition`. `editor.ChartAction` joins `WorkbookAction` / `WorksheetAction` / `StyleAction` as a fourth action type. Note that `WithChartTitle` sets the title text *and* its visibility, because text alone does not survive a save and reload as a visible title, and that `HideChartTitle` is not reversible through this API — hiding discards the text, so a reloaded file falls back to the automatic title derived from the series. Of the legend positions, `ChartLegendNotDocked` is applied but does not survive a save: the engine reports it back in memory, but XLSX cannot record an undocked legend, so a reloaded chart is docked right like any chart that never had a position set (all other positions round-trip unchanged).
- **ISSUE-CELLSGO-300**: `examples/chart` — a binding-free example that builds four chart-bearing workbooks (`datasource.NewEmptyWorkbook` seed, no `asposecells` import), the third of which is produced by re-opening the first from disk and editing chart 0 in place, so the example doubles as a demonstration that a chart written by one pass is found and changed by the next. Wired into `examples/run.sh`, `run_all.sh`, `run.ps1`, `run_all.ps1`, and the CI examples step.
- **ISSUE-CELLSGO-300**: Chart query API — `query.ChartInfo`, `query.ChartSeriesData`, and `query.ChartCount` read chart metadata, series data, and chart counts from a worksheet. `ChartMetadata` reports the chart's type, title, visibility, built-in style, legend, bounds, data range, and series count as Go-native values; `ChartSeries` reports each series' values range, category data, and data point count. The read counterpart of `editor.AddChart` and `editor.InChart`.
- **ISSUE-CELLSGO-300**: `editor.ChartStylePreset` bundles common chart configuration and styling options (chart type, built-in style, title, legend visibility, legend position, data range, category data, bounds) into a single value for reuse across charts or workbooks. `editor.WithChartPreset(preset)` applies the preset's settings in order; zero-value fields are skipped, so a preset can partially specify configuration. Enhanced from the initial 4-field version to a comprehensive 9-field preset that can completely define a chart's appearance and data.
- **ISSUE-CELLSGO-300**: Predefined chart templates — `editor.ProfessionalColumn`, `editor.MinimalPie`, `editor.PresentationBar`, `editor.DashboardLine`, `editor.ReportArea`, and `editor.SimpleScatter` are ready-to-use `ChartStylePreset` values with sensible defaults for common use cases. Use them directly for consistent styling or customize them by copying and modifying fields. Templates provide professional-looking charts with minimal code.
- **ISSUE-CELLSGO-300**: Deep chart styling control — `editor.WithChartTitleFont`, `editor.WithChartTitleColor`, `editor.WithChartLegendFont`, `editor.WithChartSeriesColor`, and `editor.WithChartSeriesName` provide fine-grained control over chart element appearance. Set title/legend fonts (name, size, bold), title color, series colors, and series names. These functions work alongside `ChartStylePreset` for comprehensive customization.

### Changed
- License configuration now reads the **`LicenseFilePath`** environment variable (the legacy `LicensePath` remains as a fallback): `core.SetLicense("")` falls back to it, every `examples/` command applies it through the new shared `examples.SetLicense()` helper, and the `tests` package applies it in `TestMain` — failing fast when a configured path is invalid instead of silently degrading to evaluation mode. `core.SetLicense` also pre-checks that a non-empty path exists and is readable, so a missing/unreadable license file is now reported as `ErrLicenseInvalid` instead of being silently ignored by the engine.
- **ISSUE-CELLSGO-279**: `datasource.DataSink.Write` now takes `(name string, data []byte) error` instead of returning an `io.WriteCloser`; `DataSource.ByteData()` (which swallowed read errors by returning nil) is removed, and all reads go through `internal/aspose/cells.ReadSource` with full error propagation. `ReaderSource` now wraps an `io.Reader` instead of an `io.ReadCloser`, and `FileStore` is gone (its two capabilities are `FilePathSource` / `FilePathSink`).
- **ISSUE-CELLSGO-279**: Removed the three byte-returning `transfer` exports whose names collide with the new sink-based signatures (`ExportWorksheetToJson`, `ExportRangeToJson`, `ExportSpreadsheetToXml`); use the sink-based versions with a `*datasource.BytesSink` instead.
- **ISSUE-CELLSGO-279**: The former variant entry points (`ConvertSpreadsheet`, `ConvertToWriter`, `ConvertSpreadsheetToFile`, `MergeSpreadsheets*`, `SplitSpreadsheet*`, `ImportCSVFile`, `ImportJsonFile`, `ImportXMLFile`, `Import*DataIntoSpreadsheet`, and the `*ToFile` exports) are kept as deprecated thin wrappers over the composites. New code should call the composite with an appropriate sink.
- **ISSUE-CELLSGO-279**: The examples now demonstrate the sink-based composites (`datasource.BytesSink`, `datasource.NewWriterSink`, `datasource.FilePathSink`, `datasource.FolderSink`, `datasource.NewZipSink`) and the transfer Options; outputs are unchanged under `examples/<name>/out`.
- **ISSUE-CELLSGO-279**: Replaced the root `main.go` smoke test with `examples/` commands decomposed by功能: `convert` shows conversion through every sink shape, `edit` demonstrates the full worksheet operation chain, `merge-split` covers merge/split through every sink shape, and `transfer` covers JSON/XML export and CSV/JSON/XML import. Each example is self-contained so it runs headlessly.
- **ISSUE-CELLSGO-279**: Moved the sample data from `TestData/Source/` into `examples/data/` and deleted `TestData/` (its `Output/` subdirectory held only generated artifacts). Doc comments and `docs/` now reference `examples/data/` paths.
- **ISSUE-CELLSGO-279**: Added `examples/run.sh` to execute every example in one go or a single example individually (`./examples/run.sh` or `./examples/run.sh convert`); CI drives the examples through it on Linux and Windows after the test suite, with retries for the native engine's intermittent splitter failure.
- **ISSUE-CELLSGO-279**: Moved the shared example helper out of `internal/` into `examples/common` so example-only code stays out of the library packages; example outputs now use cwd-independent absolute paths under `examples/<name>/out`, so `go run ./examples/...` from the module root never pollutes it.
- **ISSUE-CELLSGO-279**: The `transfer` example avoids relying on the default first sheet's name, which evaluation mode can corrupt; it puts its table on an explicitly named sheet.
- **ISSUE-CELLSGO-297**: `query` and `transfer` now share the same sheet-selection type (`internal/aspose/cells.Sheet`, index or name) with `WithSheetIndex` / `WithSheet`; both default to the first worksheet by index.
- **ISSUE-CELLSGO-299**: `TestQueryReadCellTypes` now reads the worksheet once via `ReadWorksheet` and inspects the returned grid, instead of issuing one `ReadCell` load per cell — every query call re-loads the workbook, and the evaluation copy caps total loads per process. This frees the load budget needed by the new named-range/comment/encryption tests.
- **ISSUE-CELLSGO-300**: The chart entry points are split by what they act on. `AddChart`, `DeleteChart` and `DeleteAllCharts` are worksheet operations — they change which charts a sheet has — and now live in `editor/worksheet.go` alongside the other worksheet operations. `InChart`, the `ChartAction`s and the `ChartType` / `ChartLegendPosition` types change a chart that already exists and stay in `editor/charts.go`. The docs are ordered to match.
- **ISSUE-CELLSGO-300**: The CI examples step now also runs the `query` example, which had been missing from it since the examples were added.
- **ISSUE-CELLSGO-300**: The chart type/legend-name and range/bounds validation tests moved from `internal/aspose/cells/charts_test.go` to `tests/charthelpers_test.go`, so every test file in the repository now lives in `tests/`. The move forced the tests off the internal package's unexported helpers and onto its exported resolution and validation functions, which meant the exhaustive type test could no longer read the name table: it now states all 81 names itself and asserts they resolve to 81 distinct values covering the engine's enum, which is the same bijection proof with the names independently pinned.

---

## [v26.8.0] - 2026-08-30

### Added
- **ISSUE-CELLSGO-263**: Export features — export worksheets or cell ranges to JSON / XML (`transfer.ExportRangeToJson`, `ExportWorksheetToJson`, `ExportSpreadsheetToXml`, and their `*File` variants)
- **ISSUE-CELLSGO-266**: Write / import interface — import CSV / XML / JSON data into a worksheet (`transfer.ImportCSVDataIntoSpreadsheet`, `ImportXMLDataIntoSpreadsheet`, `ImportJsonDataIntoSpreadsheet`, and their `*File` variants)
- **ISSUE-CELLSGO-279**: Add / delete / rename worksheets (`editor.WithAddWorksheet`, `editor.WithDeleteWorksheet`, `editor.WithRenameWorksheet`)
- **ISSUE-CELLSGO-279**: `errors` package with sentinel errors (`ErrSaveOptionNil`, `ErrUnsupportedFormat`, `ErrInvalidOutputPath`, `ErrInvalidSheetID`, `ErrInvalidValue`, `ErrInvalidColor`, `ErrInputIsFolder`) so failures can be classified with `errors.Is`
- **ISSUE-CELLSGO-279**: Package-level documentation for every package
- **ISSUE-CELLSGO-279**: `examples/` covering convert, edit, merge-split, and transfer
- **ISSUE-CELLSGO-279**: GitHub Actions CI (build, vet, gofmt, test) on Windows and Linux

### Changed
- Updated `go.mod` to Go 1.21 and bumped the `aspose-cells-go-cpp` dependency to v26.7.0
- **ISSUE-CELLSGO-279**: Replaced the duplicated `Open() → ReadAll → NewWorkbook_Stream` pattern with shared `internal/aspose/cells` helpers
- **ISSUE-CELLSGO-279**: Deduplicated `SplitSpreadsheet` / `SplitSpreadsheetToZipWriter` via a shared `renderWorksheetOutputs` helper
- **ISSUE-CELLSGO-279**: Removed dot-imports in `transfer`, qualifying internal helpers explicitly
- **ISSUE-CELLSGO-279**: Made the `formats` registry concurrency-safe (`Register` / `Unregister` / `Get` / `List`, sorted `List`)
- **ISSUE-CELLSGO-279**: Buffered `datasource.ReaderSource` so `ByteData()` / `Open()` are repeatable; added `NewReaderSource`
- **ISSUE-CELLSGO-279**: Deduplicated the `SetCellValue` / `SetValue` type switch into a shared `toObject` helper

---

## [v26.6.1] - 2026-06-10

### Changed
- **ISSUE-CELLSGO-260**: Updated README documentation and bumped the version badge to v26.6.1

---

## [v26.6.0] - 2026-06-07

### Added
- **ISSUE-CELLSGO-254**: Added `editor` package for workbook editing capabilities
- **ISSUE-CELLSGO-256**: Support for setting styles on cells and ranges
- **ISSUE-CELLSGO-259**: Added comprehensive documentation for core packages:
    - `converter` - Document conversion utilities
    - `datasource` - Data source management
    - `editor` - Workbook editing operations
    - `manipulator` - Document manipulation functions
    - `saveoptions` - Save options configuration

### Enhanced
- **ISSUE-CELLSGO-256**: Enhanced function descriptions across all public APIs
- **ISSUE-CELLSGO-254**: Updated README documentation

---

## [v26.4.0] - 2026-04-19

### Changed
- Updated `go.mod` dependencies to latest versions

---

## [v26.3.1] - 2026-04-05

### Added
- **ISSUE-CELLSGO-243**: Enhanced Aspose.Cells for Go via C++ Toolkits documentation
- **ISSUE-CELLSGO-245**: Added `splitter` package with document splitting functions

### Enhanced
- **ISSUE-CELLSGO-245**: Enhanced `SaveOption` interface with additional configuration options

---

## [v26.3.0] - 2026-03-28

### Added
- **ISSUE-CELLSGO-239**: Added comprehensive function comments across all packages
- **ISSUE-CELLSGO-251**: Added `merge` function for combining workbooks
- **ISSUE-CELLSGO-239**: Improved import path documentation

### Optimized
- **ISSUE-CELLSGO-251**: Optimized `convert` and `split` functions for better performance

### Changed
- **ISSUE-CELLSGO-239**: Updated README documentation
- **ISSUE-CELLSGO-239**: Removed `main.go` from the codebase
- **ISSUE-CELLSGO-239**: Updated import information in documentation

---

## [v26.2.0] - 2026-03-20

### Added
- **ISSUE-CELLSGO-234**: Initial development of Aspose.Cells for Go via C++ Toolkits
    - Core package structure established
    - Basic workbook manipulation capabilities
    - Foundation for document conversion and processing

### Changed
- **Initial commit**: Project repository initialized

---

## Version Tag Reference

| Version Tag | Release Date | Key Features |
|-------------|--------------|--------------|
| v26.8.0 | 2026-08-30 | Export/import features, worksheet management, sentinel errors, CI, examples |
| v26.6.1 | 2026-06-10 | README and CHANGELOG updates |
| v26.6.0 | 2026-06-07 | Editor package, style support, enhanced docs |
| v26.4.0 | 2026-04-19 | Dependency updates |
| v26.3.1 | 2026-04-05 | Splitter package, enhanced SaveOption |
| v26.3.0 | 2026-03-28 | Merge function, optimized converters |
| v26.2.0 | 2026-03-20 | Initial release |

---

## Issue Tracking

All changes are tracked under the **CELLSGO** issue prefix. For more details, please refer to the internal issue tracking system.

- **CELLSGO-234**: Initial development
- **CELLSGO-239**: Documentation and import improvements
- **CELLSGO-243**: Enhanced documentation
- **CELLSGO-245**: Splitter and SaveOption enhancements
- **CELLSGO-251**: Merge function and optimizations
- **CELLSGO-254**: Editor package and README updates
- **CELLSGO-256**: Style support and function descriptions
- **CELLSGO-259**: Package documentation (converter, datasource, editor, manipulator, saveoptions)
- **CELLSGO-260**: README and CHANGELOG updates
- **CELLSGO-263**: Export features (JSON / XML)
- **CELLSGO-266**: Write / import interface (CSV / XML / JSON)
- **CELLSGO-279**: Worksheet management and code enhancements

---

## Semantic Versioning

This project follows [Semantic Versioning](https://semver.org/):
- **Major version (v26)**: Breaking changes
- **Minor version (8,6,4,3,2)**: New features and enhancements
- **Patch version (0,1)**: Bug fixes and documentation updates

---

*This changelog is automatically maintained based on commit history. Last updated: 2026-08-30*
