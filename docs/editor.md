# editor

Package editor provides functionality for editing spreadsheet documents using a fluent, action-based Domain Specific Language (DSL).
## Functions

### EditSpreadsheet

```go
func EditSpreadsheet 
```

EditSpreadsheet is the core entry point for the spreadsheet editing DSL.

It orchestrates the entire workflow: reading source data, loading it into the Aspose.Cells engine, applying a series of user-defined actions, and finally exporting the result as a byte slice.

Parameters:

  - source: A datasource.DataSource providing access to the raw spreadsheet data. This abstracts the input source (e.g., file, HTTP request, in-memory buffer).

  - actions: A variadic list of WorkbookAction functions. These are the specific instructions (e.g., "SetStyle", "UpdateValue") that will be executed sequentially on the loaded workbook.

Returns:

  - \[]byte: The binary representation of the modified spreadsheet after all actions have been successfully applied. This data can be written directly to a file or HTTP response.

  - error: An error object if any step in the process fails (e.g., file not found, invalid format, or an action-specific error). If successful, this is nil.

Example:

data, err := EditSpreadsheet(fileSource, WithActiveSheet("Sheet1"), InWorksheet("Sheet1", SetCellValue(0, 0, "Hello World")), )

	if err != nil {
	   log.Fatal(err)
	}

### EditSpreadsheetToSink

```go
func EditSpreadsheetToSink(source datasource.DataSource, sink datasource.DataSink, actions ...WorkbookAction) error
```

EditSpreadsheetToSink applies a series of workbook actions to a spreadsheet and
writes the result to a datasource.DataSink. It is the sink-based form of
EditSpreadsheet: the caller picks the output shape (file, io.Writer, or
in-memory bytes) by choosing the sink. The result is saved in the source
workbook's original format.

Example:

```go
var sink datasource.BytesSink
err := EditSpreadsheetToSink(fileSource, &sink, InWorksheet("Sheet1", SetCellValue(0, 0, "Hello World")))
data := sink.Bytes()
```

### SetFormula

```go
func SetFormula(row, column int, formula string) WorksheetAction
```

SetFormula assigns a formula to a specific cell. The formula is evaluated when
the workbook is calculated; pair it with CalculateAll when the result needs to
be current before the workbook is saved or read back.

Example:

```go
EditSpreadsheet(fileSource,
    InWorksheet("Sheet1",
        SetCellValue(0, 0, 100),
        SetCellValue(0, 1, 200),
        SetFormula(0, 2, "=A1+B1"),
    ),
    CalculateAll(),
)
```

### CalculateAll

```go
func CalculateAll() WorkbookAction
```

CalculateAll recalculates every formula in the workbook. Place it after the
actions that set or depend on formula values so the saved workbook holds
current results.

### WriteRows

```go
func WriteRows[T any](source datasource.DataSource, sink datasource.DataSink, rows []T, opts ...WriteRowsOption) error
```

Writes a slice of `T` into a worksheet and saves the result to a sink — the
write counterpart of [`query.ReadRows`](query.md#readrows), using the same
column mapping. `T` must be a struct; each field becomes a column named by its
`excel:"name"` tag, or by its own field name when no tag is present; `excel:"-"`
skips a field. With `WithWriteHeader(true)` the column names are written as a
header row (row 0) before the data rows; otherwise data starts at row 0. Cells
outside the written block are left untouched. The workbook is saved in its
original format.

```go
type Employee struct {
    ID   int    `excel:"id"`
    Name string `excel:"name"`
}

err := editor.WriteRows(
    datasource.FilePathSource("template.xlsx"),
    datasource.FilePathSink("employees.xlsx"),
    []Employee{{ID: 1, Name: "Ada"}},
    editor.WithSheetIndex(0),
    editor.WithWriteHeader(true),
)
```

Supported field values are the types `SetCellValue` accepts: integers, unsigned
integers, floats, `string`, `bool`, and `time.Time` (written as a date).

**WriteRowsOption** (default sheet: first worksheet by index, aligned with
`query` and `transfer`):

```go
func WithWriteHeader(b bool) WriteRowsOption   // write column names as row 0
func WithSheetIndex(i int) WriteRowsOption     // target sheet by index (recommended)
func WithSheet(name string) WriteRowsOption    // target sheet by name
```

### SetCellComment

```go
func SetCellComment(row, column int, text string) WorksheetAction
```

Adds or replaces the comment note on a specific cell. An existing comment on the
cell is overwritten. Read it back with
[`query.ReadCellComment`](query.md#readcellcomment).

```go
EditSpreadsheet(fileSource,
    InWorksheet("Sheet1",
        SetCellValue(0, 0, "total"),
        SetCellComment(0, 0, "computed in step 2"),
    ),
)
```

### ClearComments

```go
func ClearComments() WorksheetAction
```

Removes every comment on the worksheet.

### DefineNamedRange

```go
func DefineNamedRange(name string, startRow, startColumn, endRow, endColumn int) WorksheetAction
```

Defines a workbook-level **named range** referring to a block of cells on the
applied worksheet. If a name with the same text already exists, its reference is
updated instead, so the action is idempotent. The stored reference is written as
`='SheetName'!$A$1:$B$2` (absolute cell refs, quoted sheet name). A start cell
below or to the right of the end cell returns `ErrInvalidRange`.

```go
EditSpreadsheet(fileSource,
    InWorksheet("Sheet1",
        SetCellValue(0, 0, 10),
        SetCellValue(1, 1, 20),
        DefineNamedRange("Scores", 0, 0, 1, 1),
    ),
)
```

Read the names back with [`query.NamedRanges`](query.md#namedranges) and the
cells with [`query.ReadNamedRange`](query.md#readnamedrange).

### Encrypt

```go
func Encrypt(password string) WorkbookAction
```

Encrypts the workbook so the **saved file requires the password to open**. It
sets the workbook's encryption password and selects strong (AES) encryption.
Loading an encrypted file back requires supplying the password at load time,
which the toolkit's loader does not yet expose — read encrypted files with a
password through the underlying engine (`LoadOptions` + `NewWorkbook_Stream`).

```go
data, err := EditSpreadsheet(fileSource,
    InWorksheet("Sheet1", SetCellValue(0, 0, "confidential")),
    Encrypt("hunter2"),
)
```

### AddChart

```go
func AddChart(chartType ChartType, dataRange string, byColumn bool, topRow, leftColumn, bottomRow, rightColumn int, actions ...ChartAction) WorksheetAction
```

Adds a chart to the applied worksheet and returns it, so the `ChartAction`s
passed as arguments apply to the chart just created. `dataRange` is a bare
A1-style range such as `"A1:B5"`, resolved against the same worksheet; a range
carrying a sheet qualifier is rejected, because a chart cannot source its data
from another worksheet through this API. `byColumn` selects how the range is read
into series — `true` reads down the columns (one series per column), `false`
reads across the rows.

Watch out for the category rule when choosing a range. With `byColumn` true the
engine also inspects the range's **first column**: when that column holds text and
the range is at least two columns wide, it is treated as the category axis rather
than as data, so `"A1:B5"` over a text column A yields **one** series, not two,
while `"A1:C5"` yields two. A numeric first column is not consumed this way, and a
single-column range never is. Pass the categories explicitly with
`WithChartCategoryData` when the automatic choice is not what you want.

The four bounds place the chart over a cell rectangle, all zero-based, with the
same rules as `WithChartBounds`. There is no default position — every chart says
where it goes, so pick a rectangle that covers no cells you still need to read.

```go
EditSpreadsheet(fileSource,
    InWorksheet("Sheet1",
        AddChart(ChartTypeColumn, "A1:C5", true, 5, 0, 20, 7,
            WithChartTitle("Quarterly sales by region"),
            WithChartStyle(7),
            WithChartLegend(true),
            WithChartLegendPosition(ChartLegendBottom),
        ),
    ),
)
```

### DeleteChart

```go
func DeleteChart(index int) WorksheetAction
```

Removes the worksheet's chart at `index`. The charts after it shift down, so
deleting index 0 from a sheet with two charts leaves the second one at index 0.
Returns `ErrChartNotFound` when no chart has that index.

### DeleteAllCharts

```go
func DeleteAllCharts() WorksheetAction
```

Removes every chart from the worksheet. It is not an error on a sheet that has no
charts.

### InChart

```go
func InChart(index int, actions ...ChartAction) WorksheetAction
```

Applies chart actions to the worksheet's **existing** chart at `index`. Charts are
indexed from zero in the order they were added. This is the entry point for
modifying a chart in a workbook that was read from a file.

`AddChart`, `DeleteChart` and `DeleteAllCharts` above are worksheet operations —
they change which charts the sheet has — and live in `editor/worksheet.go`.
`InChart` and everything from here down changes a chart that already exists, and
lives in `editor/charts.go`. The split is by what the operation acts on, so the
entry points that take a worksheet are found alongside the other worksheet
operations.

```go
EditSpreadsheet(fileSource,
    InWorksheet("Sheet1",
        InChart(0,
            WithChartType(ChartTypeBar),
            WithChartTitle("Revised"),
            WithChartStyle(12),
        ),
    ),
)
```

### WithChartBounds

```go
func WithChartBounds(topRow, leftColumn, bottomRow, rightColumn int) ChartAction
```

Positions the chart over a cell rectangle, with all four coordinates zero-based.
The rectangle must be strictly positive (equal top/bottom rows or left/right
columns are rejected as a zero-area chart), must not be reversed, and must lie
inside the worksheet grid; anything else returns `ErrInvalidRange`.

`AddChart` takes its rectangle directly as parameters and validates it the same
way, so this action is for **moving** a chart that already exists — normally the
one `InChart` selected.

### WithChartType

```go
func WithChartType(chartType ChartType) ChartAction
```

Switches the chart to a different type. The existing data is reinterpreted, so a
pie chart and a column chart can show the same series differently without
touching the data.

### WithChartTitle

```go
func WithChartTitle(text string) ChartAction
```

Sets the chart's title text **and makes it visible**. Visibility is set along with
the text because it does not survive a save and reload on its own — the engine
writes a title that reads back as hidden the next time the file is loaded.

```go
InChart(0, WithChartTitle("Q1 vs Q2"))
```

### HideChartTitle

```go
func HideChartTitle() ChartAction
```

Hides the chart's title. **This is not reversible through this API**: hiding also
discards the title text, so a reloaded file reports the automatic title derived
from the series again rather than the text you set earlier.

### WithChartStyle

```go
func WithChartStyle(style int) ChartAction
```

Applies one of the engine's built-in chart styles, which colour and format the
chart as a whole. The style number runs from 1 to 48; anything outside that range
returns `ErrInvalidChartStyle` rather than being silently clamped by the engine.

### WithChartLegend

```go
func WithChartLegend(visible bool) ChartAction
```

Shows or hides the chart's legend. A newly added chart shows its legend docked
right by default.

### WithChartLegendPosition

```go
func WithChartLegendPosition(position ChartLegendPosition) ChartAction
```

Docks the legend at one of the engine's positions — `ChartLegendBottom`,
`ChartLegendCorner`, `ChartLegendTop`, `ChartLegendRight`, `ChartLegendLeft`, or
`ChartLegendNotDocked`. It does **not** by itself show a hidden legend; pair it
with `WithChartLegend(true)` when the legend starts hidden. An unrecognized
position name returns `ErrInvalidChartPosition`.

Every position but `ChartLegendNotDocked` survives a save and reload unchanged.
`ChartLegendNotDocked` is genuinely applied — the engine reports it back on the
live handle — but an XLSX file cannot record an undocked legend, so a chart saved
with it reloads docked right, like any chart that never had a position set. Treat
it as an in-memory-only setting.

### WithChartDataRange

```go
func WithChartDataRange(dataRange string, byColumn bool) ChartAction
```

Re-points the chart at a different source range. This is the way to change a
chart's data wholesale: the series are rebuilt from the new range, so the series
follow whatever the new range holds. See `AddChart` for the first-column category
rule, which applies here too.

```go
InChart(0, WithChartDataRange("A1:C10", true))
```

### WithChartSeries

```go
func WithChartSeries(dataRange string, byColumn bool) ChartAction
```

Appends **one more** series to the chart, in addition to the ones it already has.
Use it to build a chart from several separate ranges; use `WithChartDataRange` to
replace them all.

### RemoveChartSeries

```go
func RemoveChartSeries(index int) ChartAction
```

Removes the chart's series at `index`; the series after it shift down. Returns
`ErrChartNotFound` when the chart has no series with that index.

### ClearChartSeries

```go
func ClearChartSeries() ChartAction
```

Removes every series from the chart, leaving it with its formatting and title but
no data.

### WithChartCategoryData

```go
func WithChartCategoryData(dataRange string) ChartAction
```

Sets the range the chart's category axis labels come from. Category data set this
way survives a later `WithChartDataRange` rebuild of the series.

```go
InChart(0,
    WithChartCategoryData("A2:A5"),
    WithChartDataRange("B1:C5", true),
)
```

### ChartType

```go
type ChartType string
```

Names a chart type. The package ships curated constants for the common families —
`ChartTypeColumn`, `ChartTypeColumnStacked`, `ChartTypeColumn3D`, `ChartTypeBar`,
`ChartTypeBarStacked`, `ChartTypeBar3DClustered`, `ChartTypeLine`,
`ChartTypeLineStacked`, `ChartTypeLineWithMarkers`, `ChartTypeLine3D`,
`ChartTypeArea`, `ChartTypeAreaStacked`, `ChartTypeArea3D`, `ChartTypePie`,
`ChartTypePie3D`, `ChartTypePieExploded`, `ChartTypePieBar`,
`ChartTypeDoughnut`, `ChartTypeDoughnutExploded`, `ChartTypeScatter`,
`ChartTypeBubble`, `ChartTypeRadar`, `ChartTypeRadarFilled`, `ChartTypeStock`,
`ChartTypeSurface3D`, `ChartTypeSurfaceContour`, `ChartTypeBoxWhisker`,
`ChartTypeFunnel`, `ChartTypeParetoLine`, `ChartTypeSunburst`,
`ChartTypeTreemap`, `ChartTypeWaterfall`, `ChartTypeHistogram`, and
`ChartTypeMap` — but **any** of the engine's 81 chart types is accepted by its own
name, so the ones without a constant are still reachable:

```go
AddChart(ChartType("pyramid"), "B2:B5", true, 5, 0, 20, 7)
```

Names are matched case- and punctuation-insensitively, so `ChartTypeColumn3D`,
`ChartType("Column3D")`, `ChartType("column_3d")` and `ChartType("COLUMN 3D")` are
the same type. An unrecognized name returns `ErrInvalidChartType` rather than
falling back to a default. The type is deliberately a toolkit-native string, not
the engine's enum.

### ChartLegendPosition

```go
type ChartLegendPosition string
```

Names where a chart's legend is docked; see `WithChartLegendPosition` for the
constants. Matched case- and punctuation-insensitively, like `ChartType`.

## Types

### ChartAction

```go
type ChartAction func(chart *asposecells.Chart) error
```

ChartAction represents an operation that modifies an existing chart object. These actions are used in conjunction with chart-targeting containers and entry points like InChart and AddChart. They encapsulate chart changes such as the title, the built-in style, the legend, and the source data.

### StyleAction

```go
type StyleAction func(style *asposecells.Style) error
```

StyleAction represents an operation that modifies a style object. These actions are used in conjunction with style-targeting containers like InDefaultStyle or SetStyle. They encapsulate formatting changes such as font adjustments, color modifications, and alignment settings.

### WorkbookAction

```go
type WorkbookAction func(workbook *asposecells.Workbook) error
```

WorkbookAction represents an operation that targets the entire workbook. It is the top-level building block of the spreadsheet editing DSL. Functions like EditSpreadsheet accept a variadic list of these actions, and container functions like InWorksheet return this type to scope operations at the workbook level.

### WorksheetAction

```go
type WorksheetAction func(worksheet *asposecells.Worksheet) error
```

WorksheetAction represents an operation scoped to a specific worksheet. These actions are typically passed as nested arguments to container functions such as InWorksheet. They allow developers to perform targeted manipulations (e.g., modifying cells, setting print areas) within a single sheet.

