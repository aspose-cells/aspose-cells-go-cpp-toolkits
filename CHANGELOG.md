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

### Added
- `errors.ErrWorksheetNotFound` sentinel for classifying missing-worksheet failures with `errors.Is`.
- `errors.ErrDataSourceNil` and `errors.ErrDataSinkNil` sentinels for classifying nil source/sink failures with `errors.Is`.
- `errors.ErrNoSources` sentinel for classifying empty-merge failures with `errors.Is`.
- `datasource.NewWriterSink(w io.Writer)` and `datasource.NewZipSink(zw *zip.Writer)` constructors for the writer/zip sinks.
- **ISSUE-CELLSGO-279**: Composite entry points — `converter.Convert`, `manipulator.Merge` / `manipulator.Split`, and sink-based `transfer` exports (`ExportWorksheetToJson`, `ExportRangeToJson`, `ExportSpreadsheetToXml`) and imports (`ImportCSV`, `ImportJsonData`, `ImportXMLData`) with functional Options (`WithSheet`, `WithStartCell`, `WithEndCell`, `WithXMLMap`, `WithBeginCell`, `WithConvertNumeric`, `WithSeparator`). One composite per capability: the output shape (file, `io.Writer`, bytes, folder, or zip archive) is chosen by picking a `datasource.DataSink`.

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
