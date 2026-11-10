![](https://img.shields.io/badge/aspose.cells%20for%20Go%20via%20C++%20Toolkits-v26.11.0-green?style=for-the-badge&logo=go) ![License](https://img.shields.io/github/license/aspose-cells/aspose-cells-go-cpp-toolkits?style=for-the-badge&logo=rocket&logoColor=white)
# Aspose.Cells for Go via C++ Toolkits

## Features

**Aspose.Cells for Go via C++ Toolkits** is a Go library for reading, creating, editing, converting, and exporting Excel spreadsheets with clean, Go-idiomatic APIs.

- **Convert to 32 formats** — One call turns a workbook into XLS, XLSX, XLSB, XLSM, XLTX, XLTM, CSV, TSV, TXT, ODS, DIF, DBF, SQL, XML, PDF, DOCX, PPTX, XPS, PCL, EPUB, HTML, JSON, Markdown, PNG, JPG/JPEG, SVG, BMP, EMF, GIF, or TIF/TIFF. The output shape is chosen by the sink: file, bytes, `io.Writer`, folder, or ZIP.
- **Export charts as PNG, JPEG, SVG, or PDF** — `converter.ExportChartToBytes`, `ExportChartToFile`, and `ExportChartToSink` render any chart with explicit width, height, and JPEG quality.
- **Import & export data** — Export worksheets or cell ranges to JSON, export a spreadsheet to XML (with XML maps), and import CSV, JSON, or XML data into a worksheet.
- **Read spreadsheets into Go values** — Typed cell reads (`query`): text, int, float, bool, and date values, ranges, merged regions, sheet names, dimensions, named ranges, and cell comments — plus `query.ReadRows[T]` to map a whole table into `[]struct` via `excel` tags.
- **Inspect charts, validations & conditional formats** — `query.ChartInfo`, `ChartSeriesData`, `ValidationInfoAt`, and `ConditionalFormattingInfoAt` expose what a workbook contains, with `ChartCount`, `ValidationCount`, and `ConditionalFormattingCount` for sizing loops.
- **Merge & split workbooks** — Merge multiple spreadsheets into one; split a workbook into per-sheet files (`datasource.FolderSink`), a ZIP archive (`datasource.NewZipSink`), or zip bytes (`datasource.NewZipSink` over a `&datasource.BytesSink{}`).
- **Fluent spreadsheet editing** — Set cell values and formulas, apply cell styles (font, color, alignment, wrap), clear contents or formats, merge / unmerge ranges, insert and delete rows, columns, and ranges, add / rename / delete / activate worksheets, write `[]struct` tables with `editor.WriteRows`, attach cell comments, define named ranges, and encrypt the saved workbook with `editor.Encrypt`.
- **Charts** — Add, modify, and delete charts with `editor.AddChart` / `editor.InChart` / `editor.DeleteChart`: set the title, the built-in style (1–48), the chart type, the legend and its position, the placement, the series, and the source data. All 81 chart types are reachable by name, and chart types and legend positions are toolkit-native values rather than engine enums.
- **Conditional formatting** — Seven rule families: cell value, expression, Top 10, above average, color scale, data bar, and icon set, all composable on one range.
- **Data validation** — Whole-number and decimal ranges, lists, and custom formulas, with input prompts, error alerts, and dropdown controls.
- **Classified errors** — 37 sentinel errors such as `ErrUnsupportedFormat`, `ErrWorksheetNotFound`, `ErrRangeTooLarge`, and `ErrUnsafeSinkName` work with `errors.Is`, so failures are classified instead of string-matched.
- **Go-idiomatic design** — Clean APIs, unified error handling, and a `DataSource` / `DataSink` abstraction that covers files, bytes, and streams with one signature.
- **Safe native lifecycle** — Engine access is serialized through an internal re-entrant lock, and `core.Cleanup()` releases the native library before exit — required by engine v26.10.0.

## Overview

**Aspose.Cells for Go via C++ Toolkits** is a wrapper toolkit based on [Aspose.Cells for Go via C++](https://products.aspose.com/cells/go-cpp/), designed to provide more Go-idiomatic APIs for Excel document processing. It makes Excel manipulation simpler and more efficient in Go projects.

## Background

While the official Go via C++ version of Aspose.Cells is powerful, its API design style leans towards C++, making it somewhat cumbersome to use in Go projects. This toolkit maintains all the original powerful features while providing:

- **Go-Style APIs**: Function naming and calling conventions that follow Go idioms
- **Optimized Error Handling**: Unified error handling mechanism
- **Common Features Integration**: Encapsulation of frequently used Excel operations

### Supported platforms

- Windows x64
- Linux x64

### Environments and versions

- Go 1.21 or greater
- Aspose.Cells for Go via C++ v26.10.0

## Quick Start

### Create a directory for your project and a main.go file within. Add the following code to your main.go. 

```go
package main

import (
  "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/converter"
  "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/core"
  "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
  _ "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/register" // register every output format for extension-based dispatch
  "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/markdown"
  "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/pdf"
  "log"
  "os"
)

func main() {
  os.Exit(run())
}

func run() int {
  // From engine v26.10.0 the native library must be released before exit.
  // os.Exit in main skips defers, so the defer lives here, in a function
  // that returns to main.
  defer core.Cleanup()

  if err := core.SetLicense(os.Getenv("LicenseFilePath")); err != nil {
    log.Print(err)
    return 1
  }
  if err := converter.Convert(
    datasource.FilePathSource("examples/data/BookText.xlsx"),
    pdf.New(pdf.WithOnePagePerSheet(true)),
    datasource.FilePathSink("out/output1.pdf")); err != nil {
    log.Print(err)
    return 1
  }
  if err := converter.Convert(
    datasource.FilePathSource("examples/data/BookText.xlsx"),
    markdown.New(markdown.WithClearData(true)),
    datasource.FilePathSink("out/output1.md")); err != nil {
    log.Print(err)
    return 1
  }
  return 0
}

```

> **About the `_ "…/register"` blank import**: it registers every supported output format with the `formats` registry. It is only strictly required when an entry point resolves the format from a file extension or `formats.Get` (for example `converter.ConvertSpreadsheetToFile`, or `manipulator.Merge` with a `formats.Get("xlsx")` option). Passing an option constructor directly — `pdf.New(...)`, `markdown.New(...)` — works without it. Including the blank import is harmless and recommended.

### Initialize project go.mod

```powershell
go mod init main
```

```
module github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26

go 1.21

require github.com/aspose-cells/aspose-cells-go-cpp/v26 v26.10.0

```
### Fetch the dependencies for your project.

```powershell
go mod tidy
```
### Set your PATH to point to the shared libraries in Aspose.Cells for Go via C++ in your current command shell.

```powershell
$env:Path = $env:Path+ ";${env:GOPATH}\github.com\aspose-cells\aspose-cells-go-cpp\v26@v26.10.0\lib\win_x86_64\"
```

## Continuous integration

The repository ships two equivalent pipelines:

- **GitHub Actions** — [`.github/workflows/ci.yml`](.github/workflows/ci.yml)
- **GitLab CI** — [`.gitlab-ci.yml`](.gitlab-ci.yml)

Both pipelines run `gofmt`, `go vet`, `go build`, and a compile-only test pass
on every change. The engine test suite and the race detector run when a license
is supplied as a base64-encoded secret named `ASPOSE_LICENSE_B64`; without it
those steps are skipped with a notice, so forks and public merge requests still
get the full static/compile gate. The native libraries are resolved from the Go
module cache automatically.

See [docs/ci.md](docs/ci.md) for secret setup, local equivalents, and the
coverage floor.

## Concurrency

The Aspose.Cells engine is not safe to enter from more than one goroutine at a
time. The toolkit serializes every engine call on a re-entrant per-goroutine
lock, so concurrent calls from many goroutines are **safe but do not run in
parallel**: two conversions started on different goroutines take turns, one
after the other. Locking never deadlocks across nested entry points, because the
lock is re-entrant for the goroutine that holds it.

Throughput therefore scales with **processes, not goroutines**. For a batch of
large conversions, run several processes (a worker pool, one engine per process)
rather than a goroutine fan-out in one process. Pure Go work — parsing cell
references, building a data source, writing a sink — can run concurrently
around the engine calls.

## Memory and large workbooks

- A source is loaded fully into memory. A `datasource.FilePathSource` is opened
  by path instead of being read into a Go byte slice first, which avoids one full
  copy of the file; `BytesSource` and `ReaderSource` still hold the bytes.
- `query.ReadWorksheet`, `ReadRange`, and the other grid reads materialize at
  most 2,000,000 cells per call and return `ErrRangeTooLarge` beyond that. Read a
  large sheet in row slices with `ReadRange` instead of requesting the whole
  used range.
- `converter.Convert`, `manipulator.Merge`, and `manipulator.Split` hand the
  workbook they already loaded to the save option
  (`saveoptions.WorkbookApplier`), instead of serializing it to bytes and
  reloading a second workbook just to save it. A third-party `SaveOption` that
  only implements `Apply` still works through the old path.
- Call `core.Cleanup()` once before the process exits. Engine v26.10.0 must be
  released explicitly; skipping it crashes the process after the last test or
  request has already completed.

## Reading encrypted workbooks

The editor writes an encrypted workbook, and every read path accepts the
password as a load option:

```go
encrypted, err := editor.EditSpreadsheet(src,
	editor.Encrypt("s3cret!"),
	editor.InWorksheet(0, editor.SetCellValue(0, 0, "classified")))

// Read it back (query / transfer / converter / editor each have WithPassword).
v, err := query.ReadCell(datasource.BytesSource(encrypted), "A1",
	query.WithPassword("s3cret!"))

var out datasource.BytesSink
err = converter.Convert(datasource.BytesSource(encrypted),
	formats.Get("pdf"), &out, converter.WithPassword("s3cret!"))
```

`converter.ChartExportOptions` carries its password in a `Password` field, and
`editor.WriteRows` uses `editor.WithWriteRowsPassword`.

## Migrating from v26.8.0

The conversion, merge/split, and transfer entry points were consolidated onto sink-based composites. The older entry points are kept as deprecated thin wrappers, so existing code still compiles; new code should prefer the composites.

| Old entry point | New entry point |
| --- | --- |
| `converter.ConvertSpreadsheet(source, opt) ([]byte, error)` | `converter.Convert(source, opt, &datasource.BytesSink{})` |
| `converter.ConvertToWriter(source, w, opt)` | `converter.Convert(source, opt, datasource.NewWriterSink(w))` |
| `converter.ConvertSpreadsheetToFile(in, out)` | `converter.Convert(datasource.FilePathSource(in), formats.Get(ext), datasource.FilePathSink(out))` |
| `manipulator.MergeSpreadsheets(sources, opt)` | `manipulator.Merge(sources, opt, &datasource.BytesSink{})` |
| `manipulator.SplitSpreadsheetToFolder(in, outDir)` | `manipulator.Split(datasource.FilePathSource(in), opt, datasource.FolderSink(outDir))` |
| `manipulator.SplitSpreadsheetToZipWriter(src, zw, opt)` | `manipulator.Split(src, opt, datasource.NewZipSink(zw))` |
| `transfer.ImportCSVFile(...)` | `transfer.ImportCSV(src, csv, datasource.FilePathSink(out), opts...)` |

The three byte-returning `transfer` exports (`ExportWorksheetToJson`, `ExportRangeToJson`, `ExportSpreadsheetToXml`) were replaced by sink-based versions; use `&datasource.BytesSink{}` for bytes.

> **Note on evaluation mode**: when a workbook is loaded the engine occasionally corrupts a worksheet's name (observed ~2% of loads; any sheet, not just the default first sheet — and it is not present in the saved bytes, so the same file can load clean once and corrupt later). `query` and `transfer` therefore target the first worksheet **by index** by default, and `WithSheetIndex` is the recommended way to pick a sheet — index-based lookups never read a name, so they are immune to the corruption. By-name lookups (`WithSheet`) may fail with `ErrWorksheetNotFound`, and per-sheet operations such as `manipulator.Split` can emit a garbage-named output. For name-based code that cannot tolerate the occasional spurious miss, retry the whole operation on freshly loaded input — each retry re-reads the source and reloads the workbook, and the corruption is load-time only, so a fresh load usually comes back clean. The same corruption can also hit a **cell value**: on a tiny workbook where only a couple of cells are non-empty, a read-back can be garbage far more often than the load-time rate above (measured ~1 in 3 loads), so demos that seed a few cells in memory and read them back should retry the read too.

## Deprecation schedule

The entry points below are legacy thin wrappers kept so existing code keeps compiling through the v26 line. They are scheduled for **removal in v27.0.0**, the next major version. New code must use the sink-based composites listed in [Migrating from v26.8.0](#migrating-from-v2680).

| Package | Deprecated entry points (removed in v27.0.0) |
| --- | --- |
| `converter` | `ConvertSpreadsheet`, `ConvertToWriter`, `ConvertSpreadsheetToFile` |
| `manipulator` | `MergeSpreadsheets`, `MergeSpreadsheetsToWriter`, `MergeSpreadsheetsToFile`, `SplitSpreadsheet`, `SplitSpreadsheetToZipWriter`, `SplitSpreadsheetToFolder` |
| `transfer` | `ExportWorksheetToJsonFile`, `ExportRangeToJsonFile`, `ExportSpreadsheetToXmlFile`, `ImportCSVDataIntoSpreadsheet`, `ImportCSVFile`, `ImportJsonDataIntoSpreadsheet`, `ImportJsonFile`, `ImportXMLDataIntoSpreadsheet`, `ImportXMLFile` |

Every deprecated function carries a `// Deprecated:` doc comment pointing to its replacement. If you still call any of these, migrate before upgrading to v27.

## Supported Formats

### Support file format

| **Format**                                                        | **Description**                                                                                 | **Load** | **Save** |
| :---------------------------------------------------------------- | :---------------------------------------------------------------------------------------------- | :------- | :------- |
| [XLS](https://docs.fileformat.com/spreadsheet/xls/)               | Excel 95/5.0 - 2003 Workbook.                                                                   | &radic;  | &radic;  |
| [XLSX](https://docs.fileformat.com/spreadsheet/xlsx/)             | Office Open XML SpreadsheetML Workbook or template file, with or without macros.                | &radic;  | &radic;  |
| [XLSB](https://docs.fileformat.com/spreadsheet/xlsb/)             | Excel Binary Workbook.                                                                          | &radic;  | &radic;  |
| [XLSM](https://docs.fileformat.com/spreadsheet/xlsm/)             | Excel Macro-Enabled Workbook.                                                                   | &radic;  | &radic;  |
| [XLT](https://docs.fileformat.com/spreadsheet/xlt/)               | Excel 97 - Excel 2003 Template.                                                                 | &radic;  | &radic;  |
| [XLTX](https://docs.fileformat.com/spreadsheet/xltx/)             | Excel Template.                                                                                 | &radic;  | &radic;  |
| [XLTM](https://docs.fileformat.com/spreadsheet/xltm/)             | Excel Macro-Enabled Template.                                                                   | &radic;  | &radic;  |
| [XLAM](https://docs.fileformat.com/spreadsheet/xlam/)             | An Excel Macro-Enabled Add-In file that's used to add new functions to Excel.                   |          | &radic;  |
| [CSV](https://docs.fileformat.com/spreadsheet/csv/)               | CSV (Comma Separated Value) file.                                                               | &radic;  | &radic;  |
| [TSV](https://docs.fileformat.com/spreadsheet/tsv/)               | TSV (Tab-separated values) file.                                                                | &radic;  | &radic;  |
| [TXT](https://docs.fileformat.com/word-processing/txt/)           | Delimited plain text file.                                                                      | &radic;  | &radic;  |
| [HTML](https://docs.fileformat.com/web/html/)                     | HTML format.                                                                                    | &radic;  | &radic;  |
| [MHTML](https://docs.fileformat.com/web/mhtml/)                   | MHTML file.                                                                                     | &radic;  | &radic;  |
| [ODS](https://docs.fileformat.com/spreadsheet/ods/)               | ODS (OpenDocument Spreadsheet).                                                                 | &radic;  | &radic;  |
| [JSON](https://docs.fileformat.com/web/json/)                     | JavaScript Object Notation                                                                      | &radic;  | &radic;  |
| [DIF](https://docs.fileformat.com/spreadsheet/dif/)               | Data Interchange Format.                                                                        |          | &radic;  |
| [PDF](https://docs.fileformat.com/pdf/)                           | Adobe Portable Document Format.                                                                 |          | &radic;  |
| [XPS](https://docs.fileformat.com/page-description-language/xps/) | XML Paper Specification Format.                                                                 |          | &radic;  |
| [SVG](https://docs.fileformat.com/page-description-language/svg/) | Scalable Vector Graphics Format.                                                                |          | &radic;  |
| [TIFF](https://docs.fileformat.com/image/tiff/)                   | Tagged Image File Format                                                                        |          | &radic;  |
| [PNG](https://docs.fileformat.com/image/png/)                     | Portable Network Graphics Format                                                                |          | &radic;  |
| [BMP](https://docs.fileformat.com/image/bmp/)                     | Bitmap Image Format                                                                             |          | &radic;  |
| [EMF](https://docs.fileformat.com/image/emf/)                     | Enhanced metafile Format                                                                        |          | &radic;  |
| [JPEG](https://docs.fileformat.com/image/jpeg/)                   | JPEG is a type of image format that is saved using the method of lossy compression.             |          | &radic;  |
| [GIF](https://docs.fileformat.com/image/gif/)                     | Graphical Interchange Format                                                                    |          | &radic;  |
| [MARKDOWN](https://docs.fileformat.com/word-processing/md/)       | Represents a markdown document.                                                                 |          | &radic;  |
| [SXC](https://docs.fileformat.com/spreadsheet/sxc/)               | An XML based format used by OpenOffice and StarOffice                                           | &radic;  | &radic;  |
| [FODS](https://docs.fileformat.com/spreadsheet/fods/)             | This is an Open Document format stored as flat XML.                                             | &radic;  | &radic;  |
| [DOCX](https://docs.fileformat.com/word-processing/docx/)         | A well-known format for Microsoft Word documents that is a combination of XML and binary files. |          | &radic;  |
| [PPTX](https://docs.fileformat.com/presentation/pptx/)            | The PPTX format is based on the Microsoft PowerPoint open XML presentation file format.         |          | &radic;  |

> **Format registry:** the table above lists what the engine can *save*. The [formats registry](#) exposes a subset of those formats as extension strings that `formats.Get` accepts: `bmp`, `csv`, `dbf`, `dif`, `docx`, `emf`, `epub`, `gif`, `html`, `jpeg`, `jpg`, `json`, `md`, `ods`, `pcl`, `pdf`, `png`, `pptx`, `sql`, `svg`, `tif`, `tiff`, `tsv`, `txt`, `xls`, `xlsb`, `xlsm`, `xlsx`, `xltm`, `xltx`, `xml`, `xps`. Formats the engine can save but the registry does not expose (`xlt`, `xlam`, `mhtml`, `sxc`, `fods`) are reachable only through a custom `saveoptions.SaveOption`.

## Evaluate Aspose.Cells for Go via C++ Toolkits

You can use Aspose.Cells for Go via C++ Toolkits free of cost for evaluation.The evaluation version provides almost all functionality of the product with certain limitations. The same evaluation version becomes licensed when you purchase a license and add a couple of lines of code to apply the license.
If you want to test Aspose.Cells for Go via C++ without evaluation version limitations, you can also try a 30-Day Temporary License. Please refer to <a href="https://purchase.aspose.com/temporary-license/"> How to get a Temporary License</a>?

## Limitations of Evaluation version

The evaluation version of Aspose. Cells for Go Toolset provides complete product functionality. An evaluation watermark will be inserted when saving the file. And the evaluation version can open up to 100 files.

## Run Aspose.Cells for Go via C++ Toolkits in production

A commercial license key is required in a production environment. Please contact us to <a href="https://purchase.aspose.com/buy">purchase a commercial license</a> if you want to publish application to the product server.

