# internal/aspose/cells

Package cells provides shared helpers over the Aspose.Cells engine used by the toolkit's public packages.

It centralizes reading a `datasource.DataSource` into bytes, building a `Workbook` from it, serializing a `Workbook` back to bytes, and resolving cells and cell areas, so the `converter`, `editor`, `manipulator`, and `transfer` packages do not each re-implement the same engine plumbing.

## Overview

The `internal/aspose/cells` package serves as a bridge between the toolkit's clean APIs and the underlying Aspose.Cells for Go via C++ binding. It provides:

1. **Input/output helpers**: Reading from `DataSource` and writing `Workbook` to bytes
2. **Workaround helpers**: Handling known engine quirks (e.g., evaluation mode)
3. **Utility functions**: Common operations like resolving sheet references and cell areas

## Functions

### ReadSource

```go
func ReadSource(source datasource.DataSource) ([]byte, error)
```

Reads all bytes from a `datasource.DataSource`. It opens the source, drains it fully into memory, and closes it.

This is the single place where a `DataSource` is turned into raw bytes, so every package reads sources the same way.

**Parameters**:
- `source`: A `DataSource` providing the spreadsheet data

**Returns**:
- `[]byte`: The raw bytes of the spreadsheet
- `error`: An error if the source cannot be opened or read

### GetWorkbookWithDataSource

```go
func GetWorkbookWithDataSource(source datasource.DataSource) (*asposecells.Workbook, error)
```

Creates a workbook from a data source by reading it into bytes and calling `NewWorkbook_Stream`.

**Parameters**:
- `source`: A `DataSource` providing the spreadsheet data

**Returns**:
- `*asposecells.Workbook`: The loaded workbook
- `error`: An error if the workbook cannot be loaded

### WorkbookToByteData

```go
func WorkbookToByteData(workbook *asposecells.Workbook) ([]byte, error)
```

Serializes a workbook to bytes in its original format.

**Parameters**:
- `workbook`: The workbook to serialize

**Returns**:
- `[]byte`: The serialized workbook bytes
- `error`: An error if the workbook cannot be saved

### LoadStable

```go
func LoadStable(source datasource.DataSource, attempts int, verify func(*asposecells.Workbook) error) (*asposecells.Workbook, error)
```

Loads a workbook from source, retrying on fresh loads when the engine corrupts state at load time.

**Parameters**:
- `source`: A `DataSource` providing the spreadsheet data
- `attempts`: Maximum number of load attempts
- `verify`: A function to verify the loaded workbook is valid. Return `nil` for success, or an error to retry.

**Returns**:
- `*asposecells.Workbook`: A verified workbook
- `error`: An error if all attempts fail

**Notes**:
- In evaluation mode, `NewWorkbook_Stream` occasionally returns a workbook whose in-memory state (a worksheet name or a cell value) is garbage (~2% of loads, non-deterministic, and not present in the bytes).
- `LoadStable` reopens the source and retries up to `attempts` times, calling `verify` on each load.
- Every rejected workbook is disposed to prevent native handle leaks.

**Example**:

```go
wb, err := cells.LoadStable(source, 3, func(wb *asposecells.Workbook) error {
    // Verify the first sheet name is valid
    wss, _ := wb.GetWorksheets()
    first, _ := wss.Get_Int(0)
    name, _ := first.GetName()
    
    // Check for garbage bytes in the name
    for i := 0; i < len(name); i++ {
        if name[i] == 0 || name[i] > 127 {
            return fmt.Errorf("invalid sheet name")
        }
    }
    return nil
})
```

## Cell reference helpers

### GetCell

```go
func GetCell(ws *asposecells.Worksheet, row, col int32) (*asposecells.Cell, error)
```

Gets a cell from a worksheet by row and column index.

**Parameters**:
- `ws`: The worksheet
- `row`: Zero-based row index
- `col`: Zero-based column index

**Returns**:
- `*asposecells.Cell`: The cell
- `error`: An error if the cell cannot be accessed

### UsedRange

```go
func UsedRange(ws *asposecells.Worksheet) (rows, cols int32, err error)
```

Returns the dimensions of the worksheet's used range.

**Parameters**:
- `ws`: The worksheet

**Returns**:
- `rows`: Number of rows in the used range
- `cols`: Number of columns in the used range
- `err`: An error if the range cannot be determined

### MergedAreas

```go
func MergedAreas(ws *asposecells.Worksheet) ([]Area, error)
```

Returns the worksheet's merged cell regions.

**Parameters**:
- `ws`: The worksheet

**Returns**:
- `[]Area`: The merged regions
- `err`: An error if the merged areas cannot be retrieved

### SheetNames

```go
func SheetNames(wb *asposecells.Workbook) ([]string, error)
```

Returns the names of the workbook's worksheets in order.

**Parameters**:
- `wb`: The workbook

**Returns**:
- `[]string`: The worksheet names
- `err`: An error if the names cannot be retrieved

### FindName

```go
func FindName(wb *asposecells.Workbook, name string) (*asposecells.Name, error)
```

Finds a named range by name. Uses iteration over `NameCollection` rather than `Get_String` to avoid returning a dangling handle for non-existent names.

**Parameters**:
- `wb`: The workbook
- `name`: The name to find

**Returns**:
- `*asposecells.Name`: The named range, or `nil` if not found
- `err`: An error if the collection cannot be accessed

## Sheet resolution

### Sheet

```go
type Sheet struct {
    UseIndex bool
    Index    int
    Name     string
}
```

Represents a sheet selection, either by zero-based index or by name.

### FirstSheet

```go
var FirstSheet = Sheet{UseIndex: true, Index: 0}
```

A `Sheet` that always selects the first worksheet by index.

### Resolve

```go
func (s Sheet) Resolve(wb *asposecells.Workbook) (*asposecells.Worksheet, error)
```

Resolves the sheet selection to a worksheet.

**Parameters**:
- `wb`: The workbook

**Returns**:
- `*asposecells.Worksheet`: The resolved worksheet
- `error`: An error if the sheet cannot be resolved

**Behavior**:
- If `UseIndex` is true, returns the worksheet at `Index` (zero-based)
- If `UseIndex` is false, returns the worksheet with name `Name`
- Returns `ErrInvalidSheetID` if index is out of range
- Returns `ErrWorksheetNotFound` if name doesn't exist

## Area type

### Area

```go
type Area struct {
    Start CellRef
    End   CellRef
}
```

A rectangular cell region with inclusive start and end cells.

## CellRef type

### CellRef

```go
type CellRef struct {
    Row int
    Col int
}
```

A zero-based cell coordinate.

### String

```go
func (r CellRef) String() string
```

Renders the reference in Excel style, e.g., `(Row: 2, Col: 1)` -> `"B3"`.

### AbsoluteString

```go
func (r CellRef) AbsoluteString() string
```

Renders the reference with absolute markers, e.g., `(Row: 2, Col: 1)` -> `"$B$3"`.

## Known issues and workarounds

### Evaluation mode worksheet name corruption

**Problem**: In evaluation mode, `NewWorkbook_Stream` occasionally returns a workbook whose in-memory state (a worksheet name or a cell value) is garbage (~2% of loads, non-deterministic).

**Symptoms**:
- A worksheet name may contain null bytes or non-printable characters
- Cell values may appear as garbage bytes
- The issue is not present in the saved bytes (reloading may yield different results)

**Solution**: Use `LoadStable` with a verification function that checks for garbage values.

**Affected operations**:
- Any operation that reads worksheet names
- Any operation that reads cell values

**Mitigation**:
- Prefer `WithSheetIndex` over `WithSheet` in query/transfer operations
- Use `LoadStable` for critical operations
- Retry operations on fresh loads if you encounter intermittent failures

### Comment handle dangling

**Problem**: `Cell.GetComment()` returns a non-nil dangling handle for cells without comments, causing a crash when `GetNote()` is called.

**Solution**: Iterate `CommentCollection` instead and match by row/column index. Use `CommentCollection.Get_Int_Int` to safely check if a comment exists.

### Name collection dangling handles

**Problem**: `NameCollection.Get_String(name)` returns `err=nil` with a dangling handle for non-existent names.

**Solution**: Iterate `NameCollection` and compare names manually to avoid dangling handles.

## Best practices

1. **Use LoadStable for critical operations**: When reading worksheet names or cell values that must be correct, use `LoadStable` to handle evaluation mode corruption.

2. **Prefer index-based sheet selection**: Use `WithSheetIndex` instead of `WithSheet` to avoid name-based lookups that may fail due to evaluation mode corruption.

3. **Dispose rejected workbooks**: When using `LoadStable`, dispose rejected workbooks in the verify function to prevent native handle leaks.

4. **Handle dangling comment handles**: Never call `Cell.GetComment().GetNote()` directly. Use the `CommentCollection` iteration approach instead.

5. **Handle missing names safely**: When looking up named ranges, iterate the collection instead of using `Get_String`.