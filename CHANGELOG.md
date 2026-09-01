# Changelog

All notable changes to this project are documented in this file.

## [Unreleased]

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

### Changed
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
