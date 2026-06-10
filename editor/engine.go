package editor

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
	"io"
)

// EditSpreadsheet is the core entry point for the spreadsheet editing DSL.
//
// It orchestrates the entire workflow: reading source data, loading it into
// the Aspose.Cells engine, applying a series of user-defined actions, and
// finally exporting the result as a byte slice.
//
// Parameters:
//
//   - source: A datasource.DataSource providing access to the raw spreadsheet data.
//     This abstracts the input source (e.g., file, HTTP request, in-memory buffer).
//
//   - actions: A variadic list of WorkbookAction functions. These are the
//     specific instructions (e.g., "SetCellStyle", "UpdateValue") that will
//     be executed sequentially on the loaded workbook.
//
// Returns:
//
//   - []byte: The binary representation of the modified spreadsheet after
//     all actions have been successfully applied. This data can be written
//     directly to a file or HTTP response.
//
//   - error: An error object if any step in the process fails (e.g., file not found,
//     invalid format, or an action-specific error). If successful, this is nil.
//
// Example:
//
// data, err := EditSpreadsheet(fileSource,
// WithActiveSheet(0),
// InSheet("Sheet1", SetCellValue(0, 0, "Hello World")),
// )
//
//	if err != nil {
//	   log.Fatal(err)
//	}
func EditSpreadsheet(source datasource.DataSource, actions ...WorkbookAction) ([]byte, error) {
	reader, errOpen := source.Open()
	if errOpen != nil {
		return nil, errOpen
	}
	data, errRead := io.ReadAll(reader)
	if errRead != nil {
		return nil, errRead
	}
	reader.Close()
	workbook, _ := asposecells.NewWorkbook_Stream(data)

	for _, action := range actions {
		if err := action(workbook); err != nil {
			print("----")
			print(err)
			return nil, err
		}
	}
	fileFormat, _ := workbook.GetFileFormat()
	return workbook.Save_SaveFormat(fileFormatToSaveFormat(fileFormat))
}
