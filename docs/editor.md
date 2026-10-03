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

The range must contain more than one cell and must span more than one row when
`byColumn` is true (or more than one column when `byColumn` is false). A
single-cell range is rejected because it cannot produce a series, and a range that
spans only one row (with `byColumn` true) is rejected because it would yield a
single series with a single data point, which the engine accepts but produces a
visually empty chart. The toolkit rejects these upfront rather than letting the
engine silently produce a chart with no data.

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
rule, which applies here too, and for the range validation rules (single-cell and
single-row/column ranges are rejected).

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

Style actions are strict about names. `WithFontColor` / `WithBackgroundColor` reject an unrecognized color name with `ErrInvalidColor` (rather than handing it to the engine's `Color_FromName`, which throws an uncaught C++ exception and **terminates the process**), and `WithFontUnderline` / `WithHorizontalAlignment` / `WithVerticalAlignment` reject an unrecognized style or alignment name with `ErrInvalidFontUnderline` / `ErrInvalidTextAlignment`. None of them fall back to a default: a one-letter typo used to silently reformat the cell, and now reports instead. Names are matched case- and punctuation-insensitively, so `"Light Sea Green"` and `"lightseagreen"` are the same color.

Color arguments — `WithFontColor`, `WithBackgroundColor`, `WithChartTitleColor`, `WithChartSeriesColor`, `WithDataBar`, `WithColorScale` — share one vocabulary across the whole toolkit, the same one the save options' `WithGridlineColor` takes:

| Form | Example |
| --- | --- |
| Go `color.Color` | `color.NRGBA{R: 0xFF, A: 0x80}` (half-transparent red) |
| hex string | `"#FF0000"`, `"#FF000080"` (8-digit is `RRGGBBAA`), with or without the `#` |
| color name | `"red"`, `"Light Sea Green"` |
| ARGB `int` | `0xFFFF0000` |

A Go `color.Color` reports alpha-premultiplied channels, so the alpha is divided back out on the way in: `color.NRGBA{R: 0xFF, A: 0x80}` sets a **full-strength** red at half alpha, not a half-strength one. Anything the toolkit cannot turn into a color is `ErrInvalidColor`.

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


### ChartStylePreset

```go
type ChartStylePreset struct {
    ChartType      ChartType
    Style          int
    Title          string
    ShowLegend     bool
    ApplyLegend    bool
    LegendPosition ChartLegendPosition
    DataRange      string
    ByColumn       bool
    CategoryData   string
    Bounds         ChartBounds
}

type ChartBounds struct {
    TopRow      int
    LeftColumn  int
    BottomRow   int
    RightColumn int
}
```

Bundles common chart configuration and styling options into a single value for
reuse across multiple charts or workbooks. Rather than applying `WithChartType`,
`WithChartStyle`, `WithChartTitle`, `WithChartLegend`, `WithChartLegendPosition`,
`WithChartDataRange`, `WithChartCategoryData`, and `WithChartBounds` separately,
a preset combines them.

A preset can be partial: zero-value fields are skipped when applied, so a preset
can specify only the settings it cares about and leave the rest unchanged. This
makes presets useful for both complete chart templates and targeted style bundles.

`ShowLegend` is the one field that cannot be skipped by its zero value alone: a
`bool` cannot distinguish "hide the legend" from "leave it alone", so a preset
that never mentions the legend would otherwise hide it. `ApplyLegend` resolves
that — with it set, `ShowLegend: false` means *hide*. Leave `ApplyLegend` false
when `ShowLegend` is true (that case is applied on its own) and set it only to
hide. A preset that sets neither touches no legend.

### WithChartPreset

```go
func WithChartPreset(preset ChartStylePreset) ChartAction
```

Applies a `ChartStylePreset`'s configuration and styling options to the chart.
The preset's fields are applied in order: chart type, data range, category data,
bounds, style, title, legend visibility, and legend position.

```go
preset := editor.ChartStylePreset{
    ChartType:      editor.ChartTypeColumn,
    Style:          7,
    Title:          "Quarterly sales",
    ShowLegend:     true,
    LegendPosition: editor.ChartLegendBottom,
    DataRange:      "A1:C5",
    ByColumn:       true,
    CategoryData:   "A2:A5",
    Bounds:         editor.ChartBounds{TopRow: 5, LeftColumn: 0, BottomRow: 20, RightColumn: 7},
}
editor.InChart(0, editor.WithChartPreset(preset))
```

Returns `ErrInvalidChartType` for an unknown chart type name,
`ErrInvalidChartStyle` when `preset.Style` is outside 1..48 (unless zero),
`ErrInvalidRange` for malformed data/category ranges or reversed bounds, and
`ErrInvalidChartPosition` for an unknown legend position name.

## Chart Templates

The `editor` package ships with predefined chart templates for common use cases.
Each template is a `ChartStylePreset` with sensible defaults for a particular
chart style. Use them directly or as a starting point for custom presets.

### Available Templates

```go
// ProfessionalColumn is a clean, professional column chart with a title,
// bottom legend, and built-in style 7. Suitable for business reports.
editor.ProfessionalColumn

// MinimalPie is a simple pie chart without a legend, relying on data labels.
// Suitable for presentations where space is limited.
editor.MinimalPie

// PresentationBar is a bar chart optimized for presentations: clear title,
// right legend, and built-in style 10.
editor.PresentationBar

// DashboardLine is a line chart with markers, suitable for dashboards. Style
// 12 provides good visibility on screens.
editor.DashboardLine

// ReportArea is an area chart for showing trends over time in reports. Style
// 5 provides a clean look.
editor.ReportArea

// SimpleScatter is a scatter plot without a legend, suitable for showing
// correlations.
editor.SimpleScatter
```

### Using Templates

Apply a template directly:

```go
editor.AddChart(editor.ChartTypeColumn, "A1:C5", true, 5, 0, 20, 7,
    editor.WithChartPreset(editor.ProfessionalColumn),
    editor.WithChartTitle("Quarterly sales"),
)
```

Customize a template by copying it and modifying fields:

```go
custom := editor.ProfessionalColumn
custom.Style = 15                        // Override the style
custom.LegendPosition = editor.ChartLegendTop // Move legend to top
editor.AddChart(editor.ChartTypeColumn, "A1:C5", true, 5, 0, 20, 7,
    editor.WithChartPreset(custom),
)
```

Templates provide consistent styling with minimal code, making them ideal for
applications that generate multiple charts with a standard look.

## Deep Styling Control

The `editor` package provides fine-grained control over chart element styling
beyond the built-in styles. These functions allow you to customize fonts,
colors, and other visual properties of chart titles, legends, and series.

### Title Styling

```go
// Set title font
editor.WithChartTitleFont("Arial", 14, true)  // font name, size, bold

// Set title color
editor.WithChartTitleColor("#FF0000")  // hex color or color name
```

### Legend Styling

```go
// Set legend font
editor.WithChartLegendFont("Calibri", 10, false)  // font name, size, bold
```

### Series Styling

```go
// Set series color
editor.WithChartSeriesColor(0, "#00FF00")  // series index, color

// Set series name (appears in legend)
editor.WithChartSeriesName(0, "Q1 Sales")  // series index, name
```

### Example

```go
editor.InChart(0,
    editor.WithChartTitle("Quarterly Report"),
    editor.WithChartTitleFont("Arial", 16, true),
    editor.WithChartTitleColor("#2E74B5"),
    editor.WithChartLegendFont("Calibri", 10, false),
    editor.WithChartSeriesColor(0, "#4472C4"),
    editor.WithChartSeriesColor(1, "#ED7D31"),
    editor.WithChartSeriesName(0, "Product A"),
    editor.WithChartSeriesName(1, "Product B"),
)
```

These styling functions can be combined with `ChartStylePreset` for comprehensive
chart customization. The preset provides the overall structure, and the deep
styling functions fine-tune individual elements.

## Data Validation

Data validation controls what users can enter into cells. The editor provides
a fluent API for adding, modifying, and deleting validation rules.

### AddDataValidation

```go
func AddDataValidation(cellRange string, actions ...DataValidationAction) WorksheetAction
```

AddDataValidation creates a validation rule for the specified cell range. The
range is an A1-style reference (e.g., "A1:A10"). The validation is configured
by the provided DataValidationActions.

Example:

```go
editor.InWorksheet("Sheet1",
    editor.AddDataValidation("A1:A10",
        editor.WithValidationType(editor.ValidationTypeWholeNumber),
        editor.WithValidationOperator(editor.OperatorTypeBetween),
        editor.WithValidationFormula1("1"),
        editor.WithValidationFormula2("100"),
        editor.WithValidationErrorMessage("Please enter a number between 1 and 100"),
    ),
)
```

### InValidation

```go
func InValidation(index int, actions ...DataValidationAction) WorksheetAction
```

InValidation applies data validation actions to an existing validation at the
given index. Validations are indexed from zero in the order they were added.

Example:

```go
editor.InWorksheet("Sheet1",
    editor.InValidation(0,
        editor.WithValidationErrorMessage("Updated message"),
    ),
)
```

### DeleteValidation

```go
func DeleteValidation(index int) WorksheetAction
```

DeleteValidation removes the data validation at the given index from the worksheet.

### Validation Types

The following validation types are available:

- `ValidationTypeAnyValue` - No restriction
- `ValidationTypeWholeNumber` - Integer values only
- `ValidationTypeDecimal` - Decimal values only
- `ValidationTypeList` - Value must be from a list
- `ValidationTypeDate` - Date values only
- `ValidationTypeTime` - Time values only
- `ValidationTypeTextLength` - Text length restriction
- `ValidationTypeCustom` - Custom formula validation

### Operators

Comparison operators for numeric, date, time, and text length validations:

- `OperatorTypeBetween` - Value between two bounds
- `OperatorTypeNotBetween` - Value outside two bounds
- `OperatorTypeEqual` - Value equals
- `OperatorTypeNotEqual` - Value not equals
- `OperatorTypeLessThan` - Value less than
- `OperatorTypeLessOrEqual` - Value less than or equal
- `OperatorTypeGreaterThan` - Value greater than
- `OperatorTypeGreaterOrEqual` - Value greater than or equal
- `OperatorTypeNone` - No operator (the default; for validation types that take none, and for conditional-formatting rule types other than cell-value comparisons)

These names are mapped to the engine's own `OperatorType` enum, whose ordinals are not in the order the names suggest (`None` is 6, and the two "not" operators follow it). An unrecognized name is reported rather than silently written as `None`, which for a cell-value or expression rule would quietly change what the rule matches.

### Validation Actions

- `WithValidationType(t)` - Set the validation type
- `WithValidationOperator(op)` - Set the comparison operator
- `WithValidationFormula1(f)` - Set the first formula/value
- `WithValidationFormula2(f)` - Set the second formula/value (for Between)
- `WithValidationList(items)` - Set a dropdown list of allowed values
- `WithValidationInCellDropDown(show)` - Show/hide dropdown for list validations
- `WithValidationIgnoreBlank(ignore)` - Ignore blank cells
- `WithValidationShowInput(show)` - Show input message when cell is selected
- `WithValidationShowError(show)` - Show error alert when validation fails
- `WithValidationAlertStyle(style)` - Set alert style (information/warning/stop)
- `WithValidationErrorTitle(title)` - Set error dialog title
- `WithValidationErrorMessage(msg)` - Set error dialog message
- `WithValidationInputTitle(title)` - Set input message title
- `WithValidationInputMessage(msg)` - Set input message

### Dropdown List Example

```go
editor.AddDataValidation("C2:C100",
    editor.WithValidationList([]string{"Active", "Inactive", "Pending"}),
    editor.WithValidationInCellDropDown(true),
)
```

## Conditional Formatting

Conditional formatting applies visual formatting to cells based on their values.
The editor provides a fluent API for adding, modifying, and deleting conditional
formatting rules.

### AddConditionalFormatting

```go
func AddConditionalFormatting(sheetID interface{}, cellRange string, actions ...ConditionalFormatAction) WorkbookAction
```

AddConditionalFormatting creates a conditional formatting rule for the specified
cell range on the given worksheet. The range is an A1-style reference.

Example:

```go
editor.AddConditionalFormatting("Sheet1", "A1:A10",
    editor.WithColorScale("#F8696B", "#63BE7B", nil),
)
```

### InConditionalFormatting

```go
func InConditionalFormatting(index int, actions ...ConditionalFormatAction) WorksheetAction
```

InConditionalFormatting applies actions to an existing conditional formatting
collection at the given index.

### DeleteConditionalFormatting

```go
func DeleteConditionalFormatting(index int) WorksheetAction
```

DeleteConditionalFormatting removes the conditional formatting collection at the
given index from the worksheet.

### Color Scales

Color scales apply a gradient fill to cells based on their values.

```go
// 2-color scale (red to green)
editor.WithColorScale("#F8696B", "#63BE7B", nil)

// 3-color scale (red to yellow to green)
editor.WithColorScale("#F8696B", "#63BE7B", "#FFEB84")
```

### Data Bars

Data bars display a horizontal bar in each cell proportional to the cell's value.

```go
editor.WithDataBar("#63BE7B")
```

### Icon Sets

Icon sets display icons in cells based on their values relative to thresholds.

Available icon sets:

- `IconSetArrows3`, `IconSetArrows4`, `IconSetArrows5`
- `IconSetArrowsGray3`, `IconSetArrowsGray4`, `IconSetArrowsGray5`
- `IconSetFlags3`
- `IconSetSigns3`
- `IconSetSymbols3`, `IconSetSymbols32`
- `IconSetTrafficLights31`, `IconSetTrafficLights32`, `IconSetTrafficLights4`
- `IconSetRating4`, `IconSetRating5`
- `IconSetStars3`
- `IconSetBoxes5`, `IconSetQuarters5`, `IconSetTriangles3`

Example:

```go
editor.WithIconSet(editor.IconSetTrafficLights31)
```

### Cell Value Rules

Cell value rules apply formatting when a cell's value meets a condition.

```go
editor.WithCellValueRule(editor.OperatorTypeGreaterThan, "90", "",
    editor.WithFontColor("#FF0000"),
    editor.WithFontIsBold(true),
    editor.WithBackgroundColor("#FFC7CE"),
)
```

### Expression Rules

Expression rules apply formatting when a formula evaluates to true.

```go
editor.WithExpressionRule("=A2>100",
    editor.WithBackgroundColor("#FFFF00"),
)
```

### Above Average Rule

Highlights cells above or below the average value.

```go
editor.WithAboveAverageRule(
    editor.WithFontColor("#006100"),
    editor.WithBackgroundColor("#C6EFCE"),
)
```

### Top 10 Rule

Highlights the top N or bottom N values.

```go
// Top 10 values
editor.WithTop10Rule(10, true,
    editor.WithFontColor("#9C5700"),
    editor.WithBackgroundColor("#FFEB9C"),
)

// Bottom 5 values
editor.WithTop10Rule(5, false,
    editor.WithFontColor("#9C0006"),
    editor.WithBackgroundColor("#FFC7CE"),
)
```

### Combining Rules

You can add multiple conditional formatting collections to the same worksheet,
each with its own range and rules. Each collection can have multiple conditions.

```go
editor.AddConditionalFormatting("Sheet1", "A1:A10",
    editor.WithColorScale("#F8696B", "#63BE7B", nil),
)
editor.AddConditionalFormatting("Sheet1", "B1:B10",
    editor.WithDataBar("#4472C4"),
)
editor.AddConditionalFormatting("Sheet1", "C1:C10",
    editor.WithIconSet(editor.IconSetArrows3),
)
```
