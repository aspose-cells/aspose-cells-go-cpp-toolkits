# converter

## Functions

### Convert

```go
func Convert(source datasource.DataSource, opt saveoptions.SaveOption, sink datasource.DataSink) error
```

Convert converts a spreadsheet from the given data source into the format described by `opt` and writes the result to `sink`. It is the single conversion entry point: the output shape (file, `io.Writer`, or in-memory bytes) is chosen by picking the sink, and the target format is chosen by picking the save option.

Parameters:

  - source: A data source implementing the `datasource.DataSource` interface, which provides the input spreadsheet content (e.g., from a file, in-memory buffer, HTTP URL, etc.).
  - opt: Conversion options that define the output format and behavior, implementing the `saveoptions.SaveOption` interface (e.g., PDFSaveOption, XLSXSaveOption, CSVSaveOption, etc.).
  - sink: The output destination implementing the `datasource.DataSink` interface. Use `datasource.FilePathSink` for a file, `datasource.NewWriterSink(w)` for an `io.Writer`, or a `*datasource.BytesSink` to capture the result as bytes.

Returns:

  - error: An error if the conversion fails due to reasons such as a nil source/sink/option, unreadable source, unsupported format, missing license, or failure in the underlying Aspose.Cells engine.

Example — spreadsheet to PDF file:

```go
err := converter.Convert(
    datasource.FilePathSource("examples/data/BookText.xlsx"),
    pdf.New(pdf.WithOnePagePerSheet(true)),
    datasource.FilePathSink("out/output2.pdf"))
if err != nil {
    println(err)
    return
}
```

Example — spreadsheet to `[]byte` via a BytesSink:

```go
var out datasource.BytesSink
err := converter.Convert(
    datasource.FilePathSource("examples/data/BookText.xlsx"),
    pdf.New(pdf.WithOnePagePerSheet(true)), &out)
if err != nil {
    println(err)
    return
}
os.WriteFile("out/output2.pdf", out.Bytes(), 0644)
```

## Deprecated functions

The following entry points predate the sink-based `Convert` and are kept as thin wrappers for backward compatibility. New code should call `Convert` with the appropriate sink.

### ConvertSpreadsheet

```go
// Deprecated: use Convert with a datasource.BytesSink instead.
func ConvertSpreadsheet(source datasource.DataSource, opt saveoptions.SaveOption) ([]byte, error)
```

Returns the converted content as a byte slice.

### ConvertToWriter

```go
// Deprecated: use Convert with datasource.NewWriterSink instead.
func ConvertToWriter(source datasource.DataSource, w io.Writer, opt saveoptions.SaveOption) error
```

Writes the converted content directly to `w`.

### ConvertSpreadsheetToFile

```go
// Deprecated: use Convert with a datasource.FilePathSource and
// datasource.FilePathSink instead.
func ConvertSpreadsheetToFile(inputPath string, outputPath string) error
```

Converts a spreadsheet file from `inputPath` to `outputPath`, inferring the output format from the file extension.

## Chart Export

The `converter` package exports a single chart from a workbook to an image or
PDF, addressed by its worksheet index and its index within that worksheet's
chart collection (both zero-based).

### ChartExportFormat

```go
type ChartExportFormat string

const (
    ChartExportFormatPNG  ChartExportFormat = "png"   // PNG image
    ChartExportFormatJPEG ChartExportFormat = "jpeg"  // JPEG image
    ChartExportFormatSVG  ChartExportFormat = "svg"   // SVG vector image
    ChartExportFormatPDF  ChartExportFormat = "pdf"   // PDF document
)
```

### ChartExportOptions

```go
type ChartExportOptions struct {
    Format  ChartExportFormat  // Output format (PNG, JPEG, SVG, or PDF)
    Width   int                // Exact output width in pixels, or 0
    Height  int                // Exact output height in pixels, or 0
    Quality int                // JPEG quality 1-100; 0 leaves the engine default
}
```

**Sizing.** `Width` and `Height` are either **both** a positive pixel count, or
**both** zero. Both zero uses the chart's own size. The output is exactly the
requested size — it is not letterboxed to preserve aspect ratio. Setting only
one of the two returns `ErrInvalidChartSize`, because the engine cannot size an
image from a single dimension, and guessing the other would silently produce a
size the caller did not ask for. A negative value is the same error.

**Quality.** Applies to JPEG only. Zero leaves the engine default; any value
outside 1-100 returns `ErrInvalidValue` rather than being silently clamped.

### ExportChartToSink

```go
func ExportChartToSink(src datasource.DataSource, sink datasource.DataSink, sheetIndex, chartIndex int, opts *ChartExportOptions) error
```

Exports a chart from a workbook to a data sink. The chart is identified by its worksheet index and chart index (both zero-based).

Parameters:

  - src: The data source containing the workbook
  - sink: The data sink to write the exported chart to
  - sheetIndex: Zero-based index of the worksheet containing the chart
  - chartIndex: Zero-based index of the chart to export
  - opts: Export options. If nil, defaults to PNG format

Example:

```go
err := converter.ExportChartToSink(
    datasource.FilePathSource("workbook.xlsx"),
    datasource.FilePathSink("chart.png"),
    0, 0,
    &converter.ChartExportOptions{Format: converter.ChartExportFormatPNG},
)
```

### ExportChartToBytes

```go
func ExportChartToBytes(src datasource.DataSource, sheetIndex, chartIndex int, opts *ChartExportOptions) ([]byte, error)
```

Exports a chart from a workbook and returns it as a byte slice.

Example:

```go
data, err := converter.ExportChartToBytes(
    datasource.FilePathSource("workbook.xlsx"),
    0, 0,
    &converter.ChartExportOptions{Format: converter.ChartExportFormatPNG},
)
if err != nil {
    log.Fatal(err)
}
// data contains the PNG image bytes
```

### ExportChartToFile

```go
func ExportChartToFile(src datasource.DataSource, outputPath string, sheetIndex, chartIndex int, opts *ChartExportOptions) error
```

Convenience function that exports a chart to a file.

Example — export as PNG:

```go
err := converter.ExportChartToFile(
    datasource.FilePathSource("workbook.xlsx"),
    "chart.png",
    0, 0,
    &converter.ChartExportOptions{Format: converter.ChartExportFormatPNG},
)
```

Example — export as JPEG with custom quality:

```go
err := converter.ExportChartToFile(
    datasource.FilePathSource("workbook.xlsx"),
    "chart.jpg",
    0, 0,
    &converter.ChartExportOptions{
        Format:  converter.ChartExportFormatJPEG,
        Quality: 90,
    },
)
```

Example — export as SVG:

```go
err := converter.ExportChartToFile(
    datasource.FilePathSource("workbook.xlsx"),
    "chart.svg",
    0, 0,
    &converter.ChartExportOptions{Format: converter.ChartExportFormatSVG},
)
```

Example — export as PDF:

```go
err := converter.ExportChartToFile(
    datasource.FilePathSource("workbook.xlsx"),
    "chart.pdf",
    0, 0,
    &converter.ChartExportOptions{Format: converter.ChartExportFormatPDF},
)
```

Example — export at an exact pixel size (both dimensions required):

```go
err := converter.ExportChartToFile(
    datasource.FilePathSource("workbook.xlsx"),
    "chart-1200x800.png",
    0, 0,
    &converter.ChartExportOptions{
        Format: converter.ChartExportFormatPNG,
        Width:  1200,
        Height: 800,
    },
)
```

### Supported Formats

  - **PNG**: Lossless raster image, rendered directly by the engine. The output
    is exactly `Width` × `Height` when a size is given.
  - **JPEG**: Compressed raster image with adjustable quality, rendered directly
    by the engine.
  - **SVG**: Scalable vector graphics, resolution-independent, rendered directly
    by the engine.
  - **PDF**: A real PDF document that opens in any reader. The chart is rendered
    to a PNG, embedded in a one-sheet workbook sized to the image, and that
    workbook is saved through the PDF save option. The PDF therefore holds a
    **raster image** of the chart on a chart-sized page, not vector artwork —
    the engine exposes no vector chart-to-PDF call.

### Errors

| Condition | Sentinel |
|-----------|----------|
| nil `datasource.DataSource` | `ErrDataSourceNil` |
| nil `datasource.DataSink` | `ErrDataSinkNil` |
| Sheet index out of range | `ErrInvalidSheetID` |
| Chart index out of range | `ErrChartNotFound` |
| Unsupported format | `ErrUnsupportedFormat` |
| Only one of `Width` / `Height` set, or a negative size | `ErrInvalidChartSize` |
| JPEG `Quality` outside 1-100 | `ErrInvalidValue` |
