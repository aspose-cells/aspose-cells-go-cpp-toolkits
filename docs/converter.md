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

The `converter` package provides functions to export charts from a workbook to various image formats.

### ChartExportFormat

```go
type ChartExportFormat string

const (
    ChartExportFormatPNG  ChartExportFormat = "png"   // PNG image
    ChartExportFormatJPEG ChartExportFormat = "jpeg"  // JPEG image
    ChartExportFormatSVG  ChartExportFormat = "svg"   // SVG vector image
    ChartExportFormatPDF  ChartExportFormat = "pdf"   // PDF (EMF vector format)
)
```

### ChartExportOptions

```go
type ChartExportOptions struct {
    Format  ChartExportFormat  // Output format (PNG, JPEG, SVG, or PDF)
    Width   int                // Desired width in pixels (0 = use chart default)
    Height  int                // Desired height in pixels (0 = use chart default)
    Quality int                // JPEG quality 1-100 (only applies to JPEG format)
}
```

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

Example — export as SVG with custom dimensions:

```go
err := converter.ExportChartToFile(
    datasource.FilePathSource("workbook.xlsx"),
    "chart.svg",
    0, 0,
    &converter.ChartExportOptions{
        Format: converter.ChartExportFormatSVG,
        Width:  1200,
        Height: 800,
    },
)
```

### Supported Formats

  - **PNG**: Lossless raster image format, ideal for web display and high-quality prints
  - **JPEG**: Compressed raster image format with adjustable quality, smaller file sizes
  - **SVG**: Scalable vector graphics, resolution-independent, ideal for web and print
  - **PDF**: Vector format (implemented as EMF), suitable for high-quality printing

### Errors

| Condition | Sentinel |
|-----------|----------|
| nil `datasource.DataSource` | `ErrDataSourceNil` |
| nil `datasource.DataSink` | `ErrDataSinkNil` |
| Sheet index out of range | Wrapped engine error |
| Chart index out of range | `ErrChartNotFound` |
| Unsupported format | `ErrUnsupportedFormat` |
