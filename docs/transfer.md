# transfer

Package transfer exports spreadsheet data to structured formats and imports structured data back into spreadsheets.

Exports cover XML and per-worksheet/per-range JSON; imports cover CSV, XML, and JSON data into a worksheet. Every entry point writes its result to a `datasource.DataSink`, so the caller picks the output shape (file, writer, or in-memory bytes) by choosing the sink, and configures the target sheet, cell area, and format-specific options with transfer Options.

## Export functions

### ExportWorksheetToJson

```go
func ExportWorksheetToJson(source datasource.DataSource, sink datasource.DataSink, opts ...Option) error
```

Exports a worksheet's used range as JSON. The result is written to sink; by default the whole used area of sheet "Sheet1" is exported, and `WithSheet` / `WithStartCell` / `WithEndCell` narrow the target.

Example:

```go
err := transfer.ExportWorksheetToJson(
    datasource.FilePathSource("out/seed.xlsx"),
    datasource.FilePathSink("out/sheet.json"),
    transfer.WithSheet("Data"))
```

### ExportRangeToJson

```go
func ExportRangeToJson(source datasource.DataSource, sink datasource.DataSink, opts ...Option) error
```

Exports a cell range as JSON. The range defaults to the worksheet's full used area (`WithStartCell` "A1", no end cell); use `WithStartCell` and `WithEndCell` to export a specific range.

Example:

```go
err := transfer.ExportRangeToJson(
    datasource.FilePathSource("out/seed.xlsx"),
    datasource.FilePathSink("out/range.json"),
    transfer.WithSheet("Data"), transfer.WithStartCell("A1"),
    transfer.WithEndCell("B3"))
```

### ExportSpreadsheetToXml

```go
func ExportSpreadsheetToXml(source datasource.DataSource, sink datasource.DataSink, opts ...Option) error
```

Exports the whole workbook as XML using the given XML map name (`WithXMLMap`), writing the result to sink.

Example:

```go
err := transfer.ExportSpreadsheetToXml(
    datasource.FilePathSource("out/seed.xlsx"),
    datasource.FilePathSink("out/inventory.xml"),
    transfer.WithXMLMap("InventoryMap"))
```

## Import functions

### ImportCSV

```go
func ImportCSV(source datasource.DataSource, csvData datasource.DataSource, sink datasource.DataSink, opts ...Option) error
```

Imports CSV data into a worksheet of the source workbook and writes the resulting workbook to sink. The target sheet defaults to "Sheet1", the top-left cell to (0,0); use `WithSheet` / `WithBeginCell` / `WithConvertNumeric` / `WithSeparator` to change the defaults.

Example:

```go
err := transfer.ImportCSV(
    datasource.FilePathSource("out/seed.xlsx"),
    datasource.FilePathSource("examples/data/BookCsvDuplicateData.csv"),
    datasource.FilePathSink("out/imported-csv.xlsx"),
    transfer.WithSheet("Imported"), transfer.WithBeginCell(0, 0),
    transfer.WithConvertNumeric(true), transfer.WithSeparator(","))
```

### ImportJsonData

```go
func ImportJsonData(source datasource.DataSource, jsonData datasource.DataSource, sink datasource.DataSink, opts ...Option) error
```

Imports JSON data into a worksheet of the source workbook and writes the resulting workbook to sink. The target sheet defaults to "Sheet1", the top-left cell to (0,0); use `WithSheet` / `WithBeginCell` to change them.

Example:

```go
err := transfer.ImportJsonData(
    datasource.FilePathSource("out/seed.xlsx"),
    datasource.FilePathSource("data.json"),
    datasource.FilePathSink("out/imported.xlsx"),
    transfer.WithSheet("Imported"), transfer.WithBeginCell(0, 0))
```

### ImportXMLData

```go
func ImportXMLData(source datasource.DataSource, xmlData datasource.DataSource, sink datasource.DataSink, opts ...Option) error
```

Imports XML data into a worksheet of the source workbook and writes the resulting workbook to sink. The target sheet defaults to "Sheet1", the top-left cell to (0,0); use `WithSheet` / `WithBeginCell` to change them.

Example:

```go
err := transfer.ImportXMLData(
    datasource.FilePathSource("out/seed.xlsx"),
    datasource.FilePathSource("data.xml"),
    datasource.FilePathSink("out/imported.xlsx"),
    transfer.WithSheet("Imported"), transfer.WithBeginCell(0, 0))
```

## Options

### WithSheet

```go
func WithSheet(name string) Option
```

Sets the worksheet to export from or import into, by name.

> **Note**: in evaluation mode the engine occasionally corrupts a worksheet's name when the workbook is loaded (observed ~2% of loads, any sheet, not just the default first sheet), so a name-based lookup may fail with `ErrWorksheetNotFound` even for a sheet that exists. Prefer `WithSheetIndex`, which is immune to name corruption. When targeting by name, use an explicitly named sheet — create it via `editor.WithAddWorksheet` or rename the source sheet — and callers that cannot tolerate the occasional spurious miss should retry the whole operation on fresh input.

### WithSheetIndex

```go
func WithSheetIndex(i int) Option
```

Sets the worksheet to export from or import into by its zero-based index. Index-based lookup is immune to the evaluation-mode load-time name corruption, so it is the recommended way to target a sheet. The default targets the first worksheet by index.

### WithStartCell

```go
func WithStartCell(ref string) Option
```

Sets the top-left cell of an export range, e.g. "A1". Only applicable to JSON export functions.

### WithEndCell

```go
func WithEndCell(ref string) Option
```

Sets the bottom-right cell of an export range, e.g. "B3". When unset, the range extends to the worksheet's last used cell. Only applicable to JSON export functions.

### WithXMLMap

```go
func WithXMLMap(name string) Option
```

Sets the XML map name used by `ExportSpreadsheetToXml`, e.g. "InventoryMap". The XML map must exist in the source workbook.

### WithBeginCell

```go
func WithBeginCell(row, col int) Option
```

Sets the top-left cell of an import as a (row, column) pair, both zero-based. Only applicable to import functions.

### WithConvertNumeric

```go
func WithConvertNumeric(b bool) Option
```

Controls whether numeric-looking CSV fields are imported as numbers rather than text. Defaults to `true`. Only applicable to `ImportCSV`.

### WithSeparator

```go
func WithSeparator(s string) Option
```

Sets the field separator used by `ImportCSV`. Defaults to ",". Only applicable to `ImportCSV`.

## Deprecated functions

The following entry points predate the sink-based composites and are kept as thin wrappers for backward compatibility. New code should call the sink-based functions with the appropriate sink.

| Old function | New function |
|--------------|--------------|
| `ExportWorksheetToJsonFile` | `ExportWorksheetToJson` with `FilePathSource`/`FilePathSink` |
| `ExportRangeToJsonFile` | `ExportRangeToJson` with `FilePathSource`/`FilePathSink` |
| `ExportSpreadsheetToXmlFile` | `ExportSpreadsheetToXml` with `FilePathSource`/`FilePathSink` |
| `ImportCSVDataIntoSpreadsheet` | `ImportCSV` with `FilePathSource`/`FilePathSink` |
| `ImportCSVFile` | `ImportCSV` with `FilePathSource`/`FilePathSink` |
| `ImportJsonDataIntoSpreadsheet` | `ImportJsonData` with `FilePathSource`/`FilePathSink` |
| `ImportJsonFile` | `ImportJsonData` with `FilePathSource`/`FilePathSink` |
| `ImportXMLDataIntoSpreadsheet` | `ImportXMLData` with `FilePathSource`/`FilePathSink` |
| `ImportXMLFile` | `ImportXMLData` with `FilePathSource`/`FilePathSink` |

## Relationship with query

`transfer` and `query` are complementary:

- **transfer**: exports spreadsheet data to external formats (JSON, XML) or imports external data into a spreadsheet. Uses `DataSink` for output.
- **query**: reads spreadsheet data into Go values (typed `CellValue` grids, sheet names, dimensions). Returns structured data directly.

Use `transfer` when you need to serialize/deserialize data to/from files or streams. Use `query` when you need to inspect or manipulate cell values programmatically.
