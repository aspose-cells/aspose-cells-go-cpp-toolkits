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

## Types

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

