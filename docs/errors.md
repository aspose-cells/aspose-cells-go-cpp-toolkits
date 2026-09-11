# errors

Package errors defines sentinel error values returned by the toolkit.

All errors returned by the toolkit are wrapped with `%w` so callers can classify failures with `errors.Is` / `errors.As` instead of matching on error message strings.

## Sentinel errors

### ErrSaveOptionNil

```go
var ErrSaveOptionNil = errors.New("save option is nil")
```

Returned when a nil `saveoptions.SaveOption` is passed to a conversion or merge entry point.

### ErrUnsupportedFormat

```go
var ErrUnsupportedFormat = errors.New("unsupported output format")
```

Returned when no output format is registered for the requested file extension. The offending extension is included in the wrapped message.

Example:

```go
_, err := converter.ConvertSpreadsheetToFile("in.xlsx", "out.xyz")
if errors.Is(err, toolkiterrors.ErrUnsupportedFormat) {
    // fall back to a different extension
}
```

### ErrInvalidOutputPath

```go
var ErrInvalidOutputPath = errors.New("invalid output path")
```

Returned when an output path cannot be used to infer a format, e.g. it has no file extension.

### ErrInvalidSheetID

```go
var ErrInvalidSheetID = errors.New("invalid sheet identifier")
```

Returned when a worksheet identifier is neither a string (sheet name) nor an integer (sheet index).

### ErrInvalidValue

```go
var ErrInvalidValue = errors.New("invalid value")
```

Returned when a value of an unsupported type is written to a cell or range. Also used for type mismatches in `query.ReadRows` and `editor.WriteRows`.

### ErrInvalidColor

```go
var ErrInvalidColor = errors.New("invalid color value")
```

Returned when a color is given in an unsupported type.

### ErrInputIsFolder

```go
var ErrInputIsFolder = errors.New("input path is a folder, expected a file")
```

Returned when a file operation receives a directory path where a file is required.

### ErrWorksheetNotFound

```go
var ErrWorksheetNotFound = errors.New("worksheet not found")
```

Returned when a worksheet is referenced by name but no worksheet with that name exists in the workbook.

### ErrDataSourceNil

```go
var ErrDataSourceNil = errors.New("data source is nil")
```

Returned when a nil `datasource.DataSource` is passed to a toolkit entry point.

### ErrDataSinkNil

```go
var ErrDataSinkNil = errors.New("data sink is nil")
```

Returned when a nil `datasource.DataSink` is passed to a toolkit entry point.

### ErrNoSources

```go
var ErrNoSources = errors.New("no data sources to merge")
```

Returned when a merge entry point receives no input sources, which would otherwise silently produce an empty workbook.

### ErrInvalidCellRef

```go
var ErrInvalidCellRef = errors.New("invalid cell reference")
```

Returned when a cell reference string (e.g. "B3") cannot be parsed into a cell coordinate.

### ErrInvalidRange

```go
var ErrInvalidRange = errors.New("invalid cell range")
```

Returned when a cell range has an invalid shape, e.g. a range whose start cell lies below or to the right of its end cell, or an area string with more than one ":" separator.

### ErrLicenseInvalid

```go
var ErrLicenseInvalid = errors.New("invalid license")
```

Returned when the Aspose.Cells license cannot be created or applied (e.g. the file is missing, unreadable, or invalid).

### ErrColumnNotFound

```go
var ErrColumnNotFound = errors.New("column not found in header row")
```

Returned when a struct field mapped by a row reader (`query.ReadRows`) has no matching column in the worksheet's header row, or when the header row is empty so no column mapping can be established.

### ErrNameNotFound

```go
var ErrNameNotFound = errors.New("named range not found")
```

Returned when a named range is referenced by name but no such name exists in the workbook.

## Error classification examples

### Distinguishing failure types

```go
_, err := converter.ConvertSpreadsheetToFile("in.xlsx", "out.xyz")
if errors.Is(err, toolkiterrors.ErrUnsupportedFormat) {
    // unsupported output format
} else if errors.Is(err, fs.ErrNotExist) {
    // input file is missing
} else if err != nil {
    // other error
}
```

### Checking for specific conditions

```go
grid, err := query.ReadRange(src, "A1", "C3")
if errors.Is(err, toolkiterrors.ErrInvalidRange) {
    // reversed or malformed range
} else if err != nil {
    // other error
}
```

### Handling missing worksheet by name

```go
v, err := query.ReadCell(src, "A1", query.WithSheet("Sheet1"))
if errors.Is(err, toolkiterrors.ErrWorksheetNotFound) {
    // sheet doesn't exist by that name
} else if err != nil {
    // other error
}
```

## Best practices

1. **Use `errors.Is` for classification**: Never compare error messages with `==` or `strings.Contains`. Always use `errors.Is(err, toolkiterrors.ErrXYZ)`.

2. **Preserve error chains**: When wrapping errors, use `%w` to include the original error:

   ```go
   return fmt.Errorf("read source: %w", err)
   ```

3. **Document which sentinel errors a function may return**: Each public function's godoc should list the specific sentinel errors it can return.

4. **Propagate errors as-is**: Don't unwrap and re-wrap errors unless adding context. Let `errors.Is` work through the chain.