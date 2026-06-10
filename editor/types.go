package editor

import asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"

// WorkbookAction represents an operation that targets the entire workbook.
// It is the top-level building block of the spreadsheet editing DSL.
// Functions like EditSpreadsheet accept a variadic list of these actions,
// and container functions like InWorksheet return this type to scope
// operations at the workbook level.
type WorkbookAction func(workbook *asposecells.Workbook) error

// WorksheetAction represents an operation scoped to a specific worksheet.
// These actions are typically passed as nested arguments to container functions
// such as InWorksheet or InSheet. They allow developers to perform targeted
// manipulations (e.g., modifying cells, setting print areas) within a single sheet.
type WorksheetAction func(worksheet *asposecells.Worksheet) error

// StyleAction represents an operation that modifies a style object.
// These actions are used in conjunction with style-targeting containers
// like InDefaultStyle or SetCellStyle. They encapsulate formatting changes
// such as font adjustments, color modifications, and alignment settings.
type StyleAction func(style *asposecells.Style) error

//type ChartAction func(chart *asposecells.Chart) error
