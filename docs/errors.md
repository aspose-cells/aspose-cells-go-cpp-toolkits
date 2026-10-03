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

### ErrInvalidCount

```go
var ErrInvalidCount = errors.New("invalid count")
```

Returned when a row, column, or cell count argument is zero or negative. The engine treats a zero count as a silent no-op, so the toolkit rejects it rather than returning success for a call that changed nothing.

### ErrInvalidShiftType

```go
var ErrInvalidShiftType = errors.New("invalid shift type")
```

Returned when a cell-shift direction name is not one the toolkit recognizes. Accepting an unknown name would silently perform a different destructive edit than the caller asked for.

### ErrRangeTooLarge

```go
var ErrRangeTooLarge = errors.New("range too large")
```

Returned when a read would materialize more cells than the toolkit is willing to allocate. A range is checked for size before any memory is reserved, so an in-grid but enormous range (e.g. an entire column range) is reported rather than exhausting memory.

```go
grid, err := query.ReadRange(src, "A1", "ZZ1000000")
if errors.Is(err, toolkiterrors.ErrRangeTooLarge) {
    // read the sheet in slices instead
}
```

### ErrUnsafeSinkName

```go
var ErrUnsafeSinkName = errors.New("unsafe output name")
```

Returned when a `datasource.DataSink` is asked to write under a name that would escape its own destination, e.g. an archive entry or file name containing `..` segments. The name in a multi-output sink comes from the source document's worksheet names, so it is untrusted input.

### ErrInvalidFontUnderline

```go
var ErrInvalidFontUnderline = errors.New("invalid font underline")
```

Returned when a font underline style name is not one the engine recognizes. Falling back to `None` would silently drop an underline the caller asked for.

### ErrInvalidTextAlignment

```go
var ErrInvalidTextAlignment = errors.New("invalid text alignment")
```

Returned when a text alignment name is not one the engine recognizes. Falling back to `General` would silently discard an alignment the caller asked for.

### ErrXMLMapNotFound

```go
var ErrXMLMapNotFound = errors.New("xml map not found")
```

Returned by `transfer.ExportSpreadsheetToXml` when the named XML map is not one the workbook defines, or the workbook defines none at all. The engine answers a request for a missing map with empty output and no error, so without this check the caller gets a successful-looking zero-byte write. The error names the maps the workbook does define.

### ErrXMLMapAmbiguous

```go
var ErrXMLMapAmbiguous = errors.New("xml map is ambiguous")
```

Returned by `transfer.ExportSpreadsheetToXml` when no XML map was named and the workbook defines more than one, so there is no single obvious choice. Name one with `transfer.WithXMLMap`.

### ErrChartNotFound

```go
var ErrChartNotFound = errors.New("chart not found")
```

Returned when a chart is referenced by an index that is out of range for the target worksheet.

### ErrInvalidChartType

```go
var ErrInvalidChartType = errors.New("invalid chart type")
```

Returned when a chart type name is not one the engine recognizes.

### ErrInvalidChartStyle

```go
var ErrInvalidChartStyle = errors.New("invalid chart style")
```

Returned when a chart style number falls outside the engine's supported 1..48 range.

### ErrInvalidChartPosition

```go
var ErrInvalidChartPosition = errors.New("invalid chart legend position")
```

Returned when a chart legend position name is not one the engine recognizes.

### ErrInvalidChartSize

```go
var ErrInvalidChartSize = errors.New("invalid chart export size")
```

Returned when a chart export requests an unusable pixel size: a negative width or height, or exactly one of the two set while the other is left at zero. The engine's `SetDesiredSize` requires both, so a half-specified size cannot be honored and is rejected rather than silently guessed.

### ErrPictureAddFailed

```go
var ErrPictureAddFailed = errors.New("picture could not be added")
```

Returned when an image cannot be embedded into a worksheet, e.g. the engine rejects the data as an undecodable image or hands back no picture collection.

### ErrValidationNotFound

```go
var ErrValidationNotFound = errors.New("validation not found")
```

Returned when a data validation is referenced by an index that is out of range for the target worksheet.

### ErrInvalidValidationType

```go
var ErrInvalidValidationType = errors.New("invalid validation type")
```

Returned when a validation type name is not one the engine recognizes.

### ErrInvalidOperatorType

```go
var ErrInvalidOperatorType = errors.New("invalid operator type")
```

Returned when an operator type name is not one the engine recognizes.

### ErrConditionNotFound

```go
var ErrConditionNotFound = errors.New("condition not found")
```

Returned when a conditional formatting condition is referenced by an index that is out of range.

### ErrInvalidFormatConditionType

```go
var ErrInvalidFormatConditionType = errors.New("invalid format condition type")
```

Returned when a format condition type name is not one the engine recognizes.

### ErrInvalidIconSetType

```go
var ErrInvalidIconSetType = errors.New("invalid icon set type")
```

Returned when an icon set type name is not one the engine recognizes.

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