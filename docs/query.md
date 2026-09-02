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

### ReadRows

```go
func ReadRows[T any](source datasource.DataSource, opts ...Option) ([]T, error)
```

Reads a worksheet's used range into a slice of `T`, mapping each struct field to
a column via the header row — the structured, Go-idiomatic way to turn a table
into typed values. `T` must be a struct; the write counterpart is
[`editor.WriteRows`](editor.md#writerows).

The first (header) row names the columns. Each struct field is matched against a
header cell — **case-insensitively, after trimming whitespace** — by its
`excel:"name"` tag, or by its own name when no tag is present. Columns in the
header that are not in the struct are ignored; an empty cell leaves the field at
its zero value.

```go
type Employee struct {
    ID   int       `excel:"id"`
    Name string    `excel:"name"`
    Hire time.Time `excel:"hired_on"`
}

rows, err := query.ReadRows[Employee](
    datasource.FilePathSource("employees.xlsx"),
    query.WithSheetIndex(0),
)
```

**Column mapping rules** (shared with `editor.WriteRows`):

| Tag                | Effect                                        |
|--------------------|-----------------------------------------------|
| `excel:"name"`     | Column header is `name` (match is case-insensitive) |
| `excel:"-"`        | Field is skipped entirely                     |
| *(no tag)*         | Column header is the field's own name         |

A mapped field with no matching header column returns `ErrColumnNotFound`; add a
tag to rename it or `excel:"-"` to ignore it.

**Supported field types**: `string`, all integer and unsigned sizes, `float32` /
`float64`, `bool`, and `time.Time`. A cell whose kind does not fit the field
type (e.g. text into an `int`), or a number that overflows it, returns an error
wrapped in `ErrInvalidValue`. A whole-number cell (`KindInt`) is accepted into a
`float` field, and any non-empty scalar cell is accepted into a `string` field
as its canonical text. An empty worksheet yields an empty (non-nil) slice.

### NamedRanges

```go
func NamedRanges(source datasource.DataSource, opts ...Option) ([]NamedRange, error)
```

Lists the workbook's **defined names** — named ranges and formula names — as
Go-native values:

```go
type NamedRange struct {
    Name     string // e.g. "MyRange"
    RefersTo string // raw reference text, e.g. "=Data!$A$1:$B$2"
    Area     Area   // referred cell area; zero for formula names
}
```

`Area` is resolved from the name's primary contiguous range; a formula name (one
referring to a computed value rather than a cell block) reports the zero area.
Options are accepted for API uniformity but none currently affect this
workbook-level listing.

### ReadNamedRange

```go
func ReadNamedRange(source datasource.DataSource, name string, opts ...Option) ([][]CellValue, error)
```

Reads the cells of a named range as a row-major grid, the same shape as
`ReadRange`. The name is resolved through the engine's range object, so the
range is read regardless of the current sheet selection — a named range may live
on any worksheet. A name that does not exist returns `ErrNameNotFound`.

```go
grid, err := query.ReadNamedRange(datasource.FilePathSource("data.xlsx"), "MyRange")
```

Write counterpart: [`editor.DefineNamedRange`](editor.md#definenamedrange).

### ReadCellComment

```go
func ReadCellComment(source datasource.DataSource, ref string, opts ...Option) (string, error)
```

Returns the comment note on a single cell identified by its Excel reference,
e.g. `"B3"`. A cell without a comment returns an empty string. Write counterpart:
[`editor.SetCellComment`](editor.md#setcellcomment).

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
| Named range not found              | `ErrNameNotFound`                 |

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
