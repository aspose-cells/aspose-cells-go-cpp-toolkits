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
