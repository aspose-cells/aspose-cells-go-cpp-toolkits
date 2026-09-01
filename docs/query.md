# query

Package query reads spreadsheet data into Go values — the read counterpart of
`transfer`. Where `transfer` serializes a sheet or range to a format (JSON,
XML), `query` loads a workbook and returns **typed Go-native data**: cell-value
grids, sheet names, dimensions, and merged regions. Entry points never expose
engine objects.

## Cell values

`CellValue` pairs a value with a `CellKind` so a program can switch on the type
before reading it:

| Kind          | Accessor           | Meaning                                        |
|---------------|--------------------|------------------------------------------------|
| `KindEmpty`   | `IsEmpty()`        | Empty cell, incl. non-anchor cells of a merge  |
| `KindText`    | `String()`         | String cell                                    |
| `KindInt`     | `Int()` (int64)    | Numeric cell with a whole-number value         |
| `KindFloat`   | `Float()` (float64)| Numeric cell with a non-whole value            |
| `KindBool`    | `Bool()`           | Boolean cell                                   |
| `KindDateTime`| `Time()` (time.Time)| Date / date-time cell                         |
| `KindError`   | `String()`         | Formula error (e.g. `"#DIV/0!"`)               |

Accessors return `(zero, false)` when the value's kind does not match the
requested type, so a mismatch is a normal condition, not a panic.

Type detection: a cell is classified from its engine value type; a numeric cell
whose number format is a date/time format is reported as `KindDateTime` (dates
are stored as serial numbers). A formula cell yields its **calculated** value
after the workbook has been calculated.

## Functions

### ReadCell

```go
func ReadCell(source datasource.DataSource, ref string, opts ...Option) (CellValue, error)
```

Reads a single cell identified by its Excel reference, e.g. `"B3"`.

```go
v, err := query.ReadCell(datasource.FilePathSource("data.xlsx"), "B3")
if v.Kind() == query.KindText {
    text, _ := v.String()
}
```

### ReadRange

```go
func ReadRange(source datasource.DataSource, startCell, endCell string, opts ...Option) ([][]CellValue, error)
```

Reads a rectangular block bounded by `startCell` and `endCell`, e.g. `"A1"` and
`"C3"`. The result is row-major: `grid[r][c]` is the cell at row
`start.Row+r`, column `start.Col+c`. A start cell below or to the right of the
end cell returns `ErrInvalidRange`.

### ReadWorksheet

```go
func ReadWorksheet(source datasource.DataSource, opts ...Option) ([][]CellValue, error)
```

Reads the worksheet's full used range as a row-major grid. An empty worksheet
yields an empty (length 0) slice.

### ReadMergedCells

```go
func ReadMergedCells(source datasource.DataSource, opts ...Option) ([]Area, error)
```

Returns the worksheet's merged cell regions, each reported exactly once from
its anchor.

### SheetNames

```go
func SheetNames(source datasource.DataSource, opts ...Option) ([]string, error)
```

Returns the names of the workbook's worksheets in order.

### Dimensions

```go
func Dimensions(source datasource.DataSource, opts ...Option) (rows, cols int, err error)
```

Returns the used range's dimensions. An empty worksheet yields `(0, 0)`.

## Options

```go
func WithSheet(name string) Option
func WithSheetIndex(i int) Option
func WithTrimSpace(b bool) Option
```

- **By default** a query targets the first worksheet **by index**, not by name
  (`"Sheet1"`), because evaluation mode occasionally corrupts worksheet names
  at load time — index-based lookup is immune to that.
- `WithSheetIndex` selects a sheet by its zero-based position; it is the
  recommended way to target a sheet.
- `WithSheet` selects a sheet by name; see the evaluation-mode note in the
  `transfer` package, the same caveat applies.
- `WithTrimSpace` trims leading/trailing whitespace on string cells before
  they are returned (default `false`).

## Cell references

`CellRef` is a zero-based cell coordinate and `Area` a rectangular region with
inclusive start and end cells. Both render in Excel style and parse back.

```go
ref, _ := query.ParseCellRef("B3")       // query.CellRef{Row: 2, Col: 1}
ref.String()                             // "B3"
area, _ := query.ParseArea("A1:C3")      // Start: A1, End: C3
area.String()                            // "A1:C3"
```

Invalid references return `ErrInvalidCellRef`; malformed areas return
`ErrInvalidRange`.

## Errors

| Condition                          | Sentinel                          |
|------------------------------------|-----------------------------------|
| nil `datasource.DataSource`        | `ErrDataSourceNil`                |
| Unparsable cell reference          | `ErrInvalidCellRef`               |
| Reversed / malformed range         | `ErrInvalidRange`                 |
| `WithSheet` name not found         | `ErrWorksheetNotFound`            |
| `WithSheetIndex` out of range      | `ErrInvalidSheetID`               |

All errors are wrapped with `%w` so they can be classified with `errors.Is`.

## Example

```go
rows, err := query.ReadWorksheet(
    datasource.FilePathSource("examples/data/BookText.xlsx"),
    query.WithSheetIndex(0),
)
if err != nil {
    log.Fatal(err)
}
for _, row := range rows {
    for _, v := range row {
        switch v.Kind() {
        case query.KindText:
            s, _ := v.String()
            fmt.Printf("%s\t", s)
        case query.KindInt:
            i, _ := v.Int()
            fmt.Printf("%d\t", i)
        case query.KindEmpty:
            fmt.Printf("-\t")
        default:
            fmt.Printf("?\t")
        }
    }
    fmt.Println()
}
```
