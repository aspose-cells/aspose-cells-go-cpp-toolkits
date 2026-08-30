# manipulator

## Functions

### Merge

```go
func Merge(sources []datasource.DataSource, opt saveoptions.SaveOption, sink datasource.DataSink) error
```

Merge combines multiple spreadsheets into a single workbook in the format described by `opt` and writes the result to `sink`. It is the single merge entry point: the output shape (file, `io.Writer`, or in-memory bytes) is chosen by picking the sink.

Parameters:

  - sources: The data sources implementing the `datasource.DataSource` interface, which provide the input spreadsheets content (e.g., from a file, in-memory buffer, etc.).
  - opt: Output options that define the merged workbook's format and behavior, implementing the `saveoptions.SaveOption` interface (e.g., PDFSaveOption, XLSXSaveOption, CSVSaveOption, etc.).
  - sink: The output destination implementing the `datasource.DataSink` interface. Use `datasource.FilePathSink` for a file, `datasource.NewWriterSink(w)` for an `io.Writer`, or a `*datasource.BytesSink` to capture the result as bytes.

Returns:

  - error: An error if the merge fails due to reasons such as a nil source/sink/option, unreadable source, unsupported format, missing license, or failure in the underlying Aspose.Cells engine.

Example — merge to a file:

```go
save_option := html.New(html.WithExportImagesAsBase64(true), html.WithSaveAsSingleFile(true))
mergedDataSource := []datasource.DataSource{
    datasource.FilePathSource("examples/data/BookText.xlsx"),
    datasource.FilePathSource("examples/data/EmployeeSalesSummary.xlsx"),
}
err := manipulator.Merge(mergedDataSource, save_option,
    datasource.FilePathSink("out/mergedOutput2.html"))
if err != nil {
    println(err)
    return
}
```

Example — merge to `[]byte` via a BytesSink:

```go
var out datasource.BytesSink
err := manipulator.Merge(mergedDataSource, save_option, &out)
if err != nil {
    println(err)
    return
}
os.WriteFile("out/mergedOutput2.html", out.Bytes(), 0644)
```

### Split

```go
func Split(source datasource.DataSource, opt saveoptions.SaveOption, sink datasource.DataSink) error
```

Split splits a spreadsheet by worksheet into multiple files in the output format described by `opt`, writing one `(filename, data)` pair per worksheet to `sink`. Each file is named `<sheet>.<ext>`, where `<ext>` is the output format extension.

Parameters:

  - source: A data source implementing the `datasource.DataSource` interface, which provides the input spreadsheet content (e.g., from a file, in-memory buffer, etc.).
  - opt: Output options that define the split files' format and behavior, implementing the `saveoptions.SaveOption` interface (e.g., PDFSaveOption, XLSXSaveOption, CSVSaveOption, etc.).
  - sink: The output destination implementing the `datasource.DataSink` interface. Use `datasource.FolderSink` for one file per worksheet in a folder, or `datasource.NewZipSink(zipWriter)` for one archive entry per worksheet.

Returns:

  - error: An error if the split fails due to reasons such as a nil source/sink/option, unreadable source, unsupported format, missing license, or failure in the underlying Aspose.Cells engine.

Example — split into a folder, one file per worksheet:

```go
save_option := image.New(image.WithImageType("png"))
err := manipulator.Split(datasource.FilePathSource("examples/data/BookText.xlsx"),
    save_option, datasource.FolderSink("out/sheets"))
if err != nil {
    println(err)
    return
}
```

Example — split into a zip archive:

```go
save_option := image.New(image.WithImageType("png"))
zipBuf := new(bytes.Buffer)
zipWriter := zip.NewWriter(zipBuf)
err := manipulator.Split(datasource.FilePathSource("examples/data/BookText.xlsx"),
    save_option, datasource.NewZipSink(zipWriter))
if err != nil {
    println(err)
    return
}
if err := zipWriter.Close(); err != nil {
    println(err)
    return
}
os.WriteFile("out/split.zip", zipBuf.Bytes(), 0644)
```

## Deprecated functions

The following entry points predate the sink-based `Merge` / `Split` and are kept as thin wrappers for backward compatibility. New code should call `Merge` or `Split` with the appropriate sink.

### MergeSpreadsheets

```go
// Deprecated: use Merge with a datasource.BytesSink instead.
func MergeSpreadsheets(sources []datasource.DataSource, opt saveoptions.SaveOption) ([]byte, error)
```

### MergeSpreadsheetsToWriter

```go
// Deprecated: use Merge with datasource.NewWriterSink instead.
func MergeSpreadsheetsToWriter(sources []datasource.DataSource, w io.Writer, opt saveoptions.SaveOption) error
```

### MergeSpreadsheetsToFile

```go
// Deprecated: use Merge with datasource.FilePathSource inputs and a
// datasource.FilePathSink instead.
func MergeSpreadsheetsToFile(inputPaths []string, outputPath string) error
```

### SplitSpreadsheet

```go
// Deprecated: use Split with a datasource.ZipSink over an in-memory zip.Writer.
func SplitSpreadsheet(source datasource.DataSource, outSaveOption saveoptions.SaveOption) ([]byte, error)
```

### SplitSpreadsheetToZipWriter

```go
// Deprecated: use Split with a datasource.ZipSink instead.
func SplitSpreadsheetToZipWriter(source datasource.DataSource, zipWriter *zip.Writer, outSaveOption saveoptions.SaveOption) error
```

### SplitSpreadsheetToFolder

```go
// Deprecated: use Split with a datasource.FolderSink instead.
func SplitSpreadsheetToFolder(inputPath string, outputFolder string) error
```

Note that `Split` names each file `<sheet>.<ext>`, whereas `SplitSpreadsheetToFolder` uses the source file's base name as a prefix (`<base>_<sheet>.<ext>`).
