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

  - actions: A variadic list of WorkbookAction functions. These are the specific instructions (e.g., "SetCellStyle", "UpdateValue") that will be executed sequentially on the loaded workbook.

Returns:

  - \[]byte: The binary representation of the modified spreadsheet after all actions have been successfully applied. This data can be written directly to a file or HTTP response.

  - error: An error object if any step in the process fails (e.g., file not found, invalid format, or an action-specific error). If successful, this is nil.

Example:

data, err := EditSpreadsheet(fileSource, WithActiveSheet(0), InSheet("Sheet1", SetCellValue(0, 0, "Hello World")), )

	if err != nil {
	   log.Fatal(err)
	}

## Types

### StyleAction

```go
type StyleAction func(style *asposecells.Style) error
```

StyleAction represents an operation that modifies a style object. These actions are used in conjunction with style-targeting containers like InDefaultStyle or SetCellStyle. They encapsulate formatting changes such as font adjustments, color modifications, and alignment settings.

### WorkbookAction

```go
type WorkbookAction func(workbook *asposecells.Workbook) error
```

WorkbookAction represents an operation that targets the entire workbook. It is the top-level building block of the spreadsheet editing DSL. Functions like EditSpreadsheet accept a variadic list of these actions, and container functions like InWorksheet return this type to scope operations at the workbook level.

### WorksheetAction

```go
type WorksheetAction func(worksheet *asposecells.Worksheet) error
```

WorksheetAction represents an operation scoped to a specific worksheet. These actions are typically passed as nested arguments to container functions such as InWorksheet or InSheet. They allow developers to perform targeted manipulations (e.g., modifying cells, setting print areas) within a single sheet.

