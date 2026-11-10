# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [26.11.0] - 2026-11-10

### Added

- Load options and password support across every read path:
  `query.WithPassword`, `transfer.WithPassword`, `converter.WithPassword`,
  `converter.ChartExportOptions.Password`, `editor.WithPassword` (with
  `EditSpreadsheetWithOptions` / `EditSpreadsheetToSinkWithOptions`), and
  `editor.WithWriteRowsPassword`. This closes the encrypted-write →
  encrypted-read loop started by `editor.Encrypt`.
- `datasource.FilePathSource` is opened by path (`engine.OpenWorkbookFile` and
  `NewWorkbook_String_LoadOptions`) instead of being read into a Go byte slice
  first, removing one full copy of the file. Missing-path errors still propagate
  unchanged, so `os.IsNotExist` keeps working.
- `editor.AddPicture` / `editor.WithPictureSpan` embed an image and optionally
  scale it over a cell rectangle.
- `query.ReadCellFormula` reads a cell's formula, and `query.ReadStyle` reads its
  font name, size, weight, slant, color, and number format as plain Go values.
- `saveoptions.WorkbookApplier`, implemented by every format in this module.
  `converter.Convert`, `manipulator.Merge`, `manipulator.Split`, and
  `transfer.ExportRangeToJson` now save the workbook they already loaded instead
  of serializing it and reloading a second workbook just to save it. A
  third-party `SaveOption` with only `Apply` keeps the old path.
- `datasource.MultiOutputSink`, implemented by `FolderSink` and `ZipSink`.
  `manipulator.Split` rejects a single-output sink with the new
  `errors.ErrInvalidSink` instead of silently keeping only the last worksheet.
- `errors.ErrInvalidWorkbook` for a `SaveWorkbook` handle that is not a
  workbook.
- Benchmarks for convert, merge, split, and worksheet reads, plus fuzz targets
  for `ParseCellRef` and `ParseArea`.
- README sections on concurrency, memory and large workbooks, and reading
  encrypted workbooks; package docs for the new APIs.

- Continuous integration for both GitHub Actions (`.github/workflows/ci.yml`)
  and GitLab CI (`.gitlab-ci.yml`): formatting, vet, build, compile-only tests,
  licensed engine tests, race detector, coverage artifact, and a 60% statement
  coverage floor.
- `core.IsLicensed()` so callers can distinguish a licensed run from the
  supported evaluation-mode path in which `SetLicense` returns `nil` without a
  configured path.
- Smoke tests for every save options format that previously had no `Apply`
  coverage (`dbf`, `dif`, `epub`, `md`, `ods`, `sql`, `xls`, `xlsb`, `xml`).
- Generated constructor smoke test that invokes every exported `saveoptions`
  `With*` option, plus a completeness check that fails when the generated list
  drifts from the source.
- Tests for the deprecated compatibility wrappers in `converter`,
  `manipulator`, and `transfer`, so the v26 source-compatibility window is
  exercised until the v27 removal.
- `CHANGELOG.md` and `docs/ci.md`.

### Fixed

- Native option objects that were left to the binding's GC finalizers — which
  enter the engine from the finalizer goroutine without the engine lock — are
  now released deterministically under the lock: the per-cell `Object` boxes in
  the editor's write path (`PutValue_Object` / `Range.SetValue`), the JSON
  layout options in `transfer.ImportJsonData`, and the image options in chart
  export. `GetWorkbookFromBytesWithXMLLoadOptions` additionally leaked its
  `XmlLoadOptions` on every XML load (`Disarm` without `Delete`); it is now
  freed, and the aliasing of `ToLoadOptions` (same native pointer) is
  documented so the borrowed wrapper is never double-freed.
- `converter`'s chart-to-PDF path no longer serializes the chart-page workbook
  and reloads it through `pdf.Apply`; it saves the live workbook via
  `WorkbookApplier`, like the rest of the module.
- `transfer.ExportRangeToJson` now loads through the shared source loader, so a
  `FilePathSource` is opened by path instead of being copied through the Go
  heap first.
- Stale package documentation that still claimed an encrypted workbook cannot
  be read back (`editor.Encrypt` and the internal `EncryptWorkbook`); every
  read path accepts a password now. `docs/converter.md` lists the
  `ChartExportOptions.Password` field.

### Changed

- Editor cell writes go through the engine's typed `PutValue_*` calls for
  `string`, `bool`, `float32/64`, `time.Time`, and integers that fit `int32` —
  one cgo call and no native allocation per cell. Wider integers keep the
  boxed `Object` path, which now frees the box immediately.
- `query.ReadRange` / `ReadWorksheet` allocate the grid as one flat backing
  array sliced into rows (up to a million small allocations saved at the read
  cap), and the per-cell hot loops in `query`, `editor`, and the internal cell
  helpers use the new reflection-free `engine.DerivePtr`.
- `editor.EditSpreadsheet*`, `editor.WriteRows`, and the three `transfer`
  importers save a `FilePathSink` result straight to disk — engine save into a
  temp file, then atomic rename, matching `FilePathSink.Write` semantics —
  instead of holding the serialized document in Go memory first.
- Statement coverage raised from 46.5% to 67.5%, above the new 60% CI floor,
  by exercising the nine previously untested save formats, every exported
  `saveoptions` option constructor, and all deprecated compatibility wrappers.
- README version badge corrected from `v26.8.0` to `v26.10.0`, and raised to
  `v26.11.0` for this release.
- The eight posts under `blogs/` were verified against the current API and
  corrected where stale: the `Split` multi-output-sink requirement, encrypted
  workbook read-back, real-PDF chart export, `Convert`'s variadic options,
  toolkit version and example/test counts, plus a runnable `core.Cleanup`
  pattern in the chart best-practices example.
- `docs/core.md` documents the evaluation-mode return value of `SetLicense` and
  the new `IsLicensed` function.
- `docs/design.md` status updated from "待评审" to "已实施(v26.10.0)".

### Removed

- The stale root `main.go` demo. It called three deprecated converter entry
  points and referenced a `TestData/` directory that is not part of the
  repository; `examples/convert` is the maintained example.

## [26.10.0] - 2026-10-19

### Added

- Sink-based composites `converter.Convert`, `manipulator.Merge`, and
  `manipulator.Split`, with `datasource.DataSource` / `datasource.DataSink`
  covering files, bytes, streams, folders, and ZIP archives.
- `query` package for typed cell, range, worksheet, merged-region, named-range,
  chart, validation, and conditional-format reads, including `query.ReadRows[T]`.
- `editor` chart, data-validation, conditional-formatting, comment, named-range,
  encryption, and `editor.WriteRows[T]` support.
- Extension-based format registry (`formats`, `register`) covering 32 output
  formats.
- 37 sentinel errors with `errors.Is` / `errors.As` classification.
- Re-entrant engine lock and handle disarmament so the native engine is never
  entered concurrently and GC finalizers cannot destroy engine-owned objects.

### Changed

- Conversion, merge/split, and transfer entry points consolidated onto one
  sink-based shape per capability. The older entry points remain as deprecated
  thin wrappers through v26 and are scheduled for removal in v27.0.0.
- `core.Cleanup()` is required before process exit with engine v26.10.0.

[Unreleased]: https://github.com/aspose-cells/aspose-cells-go-cpp-toolkits/compare/v26.11.0...HEAD
[26.11.0]: https://github.com/aspose-cells/aspose-cells-go-cpp-toolkits/compare/v26.10.0...v26.11.0
[26.10.0]: https://github.com/aspose-cells/aspose-cells-go-cpp-toolkits/releases/tag/v26.10.0
