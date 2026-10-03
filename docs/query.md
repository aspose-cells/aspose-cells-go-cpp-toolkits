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

A range naming more than 2,000,000 cells returns `ErrRangeTooLarge` before any
memory is reserved. The check is not redundant with the grid bounds: a range like
`"A1:ZZ1000000"` is entirely *valid* — every cell in it exists in the grid — and
would otherwise have the toolkit allocate a few hundred megabytes and let the OS
kill the process. Read a sheet that large in slices with `ReadRange`.

### ReadWorksheet

```go
func ReadWorksheet(source datasource.DataSource, opts ...Option) ([][]CellValue, error)
```

Reads the worksheet's full used range as a row-major grid. An empty worksheet
yields an empty (length 0) slice. A used range larger than 2,000,000 cells
returns `ErrRangeTooLarge` rather than being allocated; see `ReadRange`.

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

## Charts

### ChartInfo

```go
func ChartInfo(source datasource.DataSource, chartIndex int, opts ...Option) (ChartMetadata, error)
```

Reads the metadata for the worksheet's chart at `chartIndex`. Charts are indexed
from zero in the order they were added. A `chartIndex` out of range returns
`ErrChartNotFound`.

```go
type ChartMetadata struct {
    Index          int          // zero-based position in the worksheet's chart collection
    Type           string       // chart type, e.g. "column", "pie", "bar"
    Title          string       // title text; empty when no title or title is hidden
    TitleVisible   bool         // whether the title is shown
    Style          int          // built-in style number (1..48), or -1 for engine default
    ShowLegend     bool         // whether the legend is visible
    LegendPosition string       // legend position, e.g. "bottom", "right"; empty when hidden
    Bounds         ChartBounds  // chart position on the worksheet (zero-based)
    DataRange      string       // source data range in A1 notation, e.g. "A1:B5"
    SeriesCount    int          // number of data series in the chart
}

type ChartBounds struct {
    TopRow      int
    LeftColumn  int
    BottomRow   int
    RightColumn int
}
```

```go
info, err := query.ChartInfo(datasource.FilePathSource("data.xlsx"), 0)
if err != nil {
    log.Fatal(err)
}
fmt.Println(info.Type, info.Title, info.SeriesCount)
```

Write counterpart: [`editor.AddChart`](editor.md#addchart) and
[`editor.InChart`](editor.md#inchart).

### ChartSeriesData

```go
func ChartSeriesData(source datasource.DataSource, chartIndex int, opts ...Option) ([]ChartSeries, error)
```

Reads the series data for the worksheet's chart at `chartIndex`. Each series
reports its values range, category data, and data point count. A `chartIndex`
out of range returns `ErrChartNotFound`.

```go
type ChartSeries struct {
    Index          int    // zero-based position in the chart's series collection
    Values         string // cell range the series plots, e.g. "$B$2:$B$5"
    CategoryData   string // cell range for category axis labels; empty when none
    DataValueCount int    // number of data points in the series
}
```

```go
series, err := query.ChartSeriesData(datasource.FilePathSource("data.xlsx"), 0)
if err != nil {
    log.Fatal(err)
}
for _, s := range series {
    fmt.Println(s.Values, s.DataValueCount)
}
```

### ChartCount

```go
func ChartCount(source datasource.DataSource, opts ...Option) (int, error)
```

Returns the number of charts on the worksheet.

```go
n, err := query.ChartCount(datasource.FilePathSource("data.xlsx"))
if err != nil {
    log.Fatal(err)
}
fmt.Printf("worksheet has %d chart(s)\n", n)
```

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
| Chart index out of range           | `ErrChartNotFound`                |

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

## Data Validation

The `query` package provides functions to read data validation rules from a
worksheet. Validation rules control what users can enter into cells.

### ValidationCount

```go
func ValidationCount(src datasource.DataSource, opts ...Option) (int, error)
```

Returns the number of data validations on the specified worksheet.

```go
count, err := query.ValidationCount(source, query.WithSheet("Sheet1"))
if err != nil {
    log.Fatal(err)
}
fmt.Printf("Validations: %d\n", count)
```

### ValidationInfoAt

```go
func ValidationInfoAt(src datasource.DataSource, index int, opts ...Option) (ValidationInfo, error)
```

Returns metadata about the data validation at the given index on the specified
worksheet. `ValidationInfo` contains:

- `Index` - Zero-based position
- `Type` - Validation type (e.g., "wholeNumber", "list", "custom")
- `Operator` - Comparison operator (e.g., "between", "greaterThan")
- `Formula1`, `Formula2` - Validation formulas/values, read back without a leading
  `=`. The engine stores every formula `=` -prefixed and hands it back that way, so
  the toolkit strips exactly one `=` on read: `WithValidationFormula1("1")` reads
  back as `"1"`, not `"=1"`. The one exception is a **list** validation, whose
  `Formula1` holds literal comma-separated values rather than a formula; the engine
  does not prefix it and the toolkit does not strip it, so a list of
  `["Yes", "No"]` reads back as `"Yes,No"`.
- `Areas` - Cell ranges the validation applies to
- `ErrorMessage`, `ErrorTitle` - Error alert text
- `InputMessage`, `InputTitle` - Input prompt text
- `ShowError`, `ShowInput` - Whether alerts/prompts are shown
- `IgnoreBlank` - Whether blank cells are ignored
- `InCellDropDown` - Whether dropdown is shown for list validations

```go
info, err := query.ValidationInfoAt(source, 0, query.WithSheet("Sheet1"))
if err != nil {
    log.Fatal(err)
}
fmt.Printf("Type: %s, Operator: %s\n", info.Type, info.Operator)
fmt.Printf("Formula1: %s, Formula2: %s\n", info.Formula1, info.Formula2)
if info.ErrorMessage != "" {
    fmt.Printf("Error: %s\n", info.ErrorMessage)
}
```

### AllValidations

```go
func AllValidations(src datasource.DataSource, opts ...Option) ([]ValidationInfo, error)
```

Returns metadata about all data validations on the specified worksheet.

```go
validations, err := query.AllValidations(source, query.WithSheet("Sheet1"))
if err != nil {
    log.Fatal(err)
}
for _, v := range validations {
    fmt.Printf("Validation %d: Type=%s, Areas=%v\n", v.Index, v.Type, v.Areas)
}
```

## Conditional Formatting

The `query` package provides functions to read conditional formatting rules from
a worksheet. Conditional formatting applies visual formatting based on cell values.

### ConditionalFormattingCount

```go
func ConditionalFormattingCount(src datasource.DataSource, opts ...Option) (int, error)
```

Returns the number of conditional formatting collections on the specified worksheet.

```go
count, err := query.ConditionalFormattingCount(source, query.WithSheet("Sheet1"))
if err != nil {
    log.Fatal(err)
}
fmt.Printf("Conditional formattings: %d\n", count)
```

### ConditionalFormattingInfoAt

```go
func ConditionalFormattingInfoAt(src datasource.DataSource, index int, opts ...Option) (ConditionalFormattingInfo, error)
```

Returns metadata about the conditional formatting collection at the given index.
`ConditionalFormattingInfo` contains:

- `Index` - Zero-based position
- `Areas` - Cell ranges the formatting applies to
- `Conditions` - List of conditions in this collection

Each `ConditionInfo` contains:

- `Index` - Zero-based position in the collection
- `Type` - Condition type (e.g., "colorScale", "dataBar", "iconSet", "cellValue")
- `Operator` - Comparison operator
- `Formula1`, `Formula2` - Condition formulas/values, read back without the
  leading `=` the engine stores them with — a rule set as `"90"` reads back as
  `"90"`, and one set as `"=A2>80"` reads back as `"A2>80"`. Conditions carry no
  list form, so the prefix is always stripped. `Formula1`/`Formula2` are empty for
  `colorScale`, `dataBar`, and `iconSet` conditions, which hold no formula.

```go
info, err := query.ConditionalFormattingInfoAt(source, 0, query.WithSheet("Sheet1"))
if err != nil {
    log.Fatal(err)
}
fmt.Printf("Areas: %v\n", info.Areas)
for _, cond := range info.Conditions {
    fmt.Printf("  Condition %d: Type=%s, Operator=%s\n", cond.Index, cond.Type, cond.Operator)
}
```

### AllConditionalFormattings

```go
func AllConditionalFormattings(src datasource.DataSource, opts ...Option) ([]ConditionalFormattingInfo, error)
```

Returns metadata about all conditional formatting collections on the specified
worksheet.

```go
formattings, err := query.AllConditionalFormattings(source, query.WithSheet("Sheet1"))
if err != nil {
    log.Fatal(err)
}
for _, cf := range formattings {
    fmt.Printf("Conditional Formatting %d: Areas=%v\n", cf.Index, cf.Areas)
    for _, cond := range cf.Conditions {
        fmt.Printf("  Condition: Type=%s\n", cond.Type)
    }
}
```
