package editor

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
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
//     specific instructions (e.g., "SetStyle", "UpdateValue") that will
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
//
//	WithActiveSheet("Sheet1"),
//	InWorksheet("Sheet1", SetCellValue(0, 0, "Hello World")),
//
// )
//
//	if err != nil {
//	   log.Fatal(err)
//	}
func EditSpreadsheet(source datasource.DataSource, actions ...WorkbookAction) ([]byte, error) {
	var sink datasource.BytesSink
	if err := EditSpreadsheetToSink(source, &sink, actions...); err != nil {
		return nil, err
	}
	return sink.Bytes(), nil
}

// EditSpreadsheetToSink applies a series of workbook actions to a spreadsheet
// and writes the result to a datasource.DataSink. It is the sink-based form of
// EditSpreadsheet: the caller picks the output shape (file, io.Writer, or
// in-memory bytes) by choosing the sink.
//
// The result is saved in the source workbook's original format.
//
// Parameters:
//
//   - source: A datasource.DataSource providing access to the raw spreadsheet data.
//
//   - sink: A datasource.DataSink that receives the modified spreadsheet.
//     A datasource.BytesSink can be used when the result is needed as bytes.
//
//   - actions: A variadic list of WorkbookAction functions applied sequentially
//     to the loaded workbook.
//
// Returns:
//
//   - error: An error object if any step in the process fails (e.g., file not
//     found, invalid format, or an action-specific error). If successful, this
//     is nil.
func EditSpreadsheetToSink(source datasource.DataSource, sink datasource.DataSink, actions ...WorkbookAction) error {
	if source == nil {
		return toolkiterrors.ErrDataSourceNil
	}
	if sink == nil {
		return toolkiterrors.ErrDataSinkNil
	}
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return err
	}
	for _, action := range actions {
		if err := action(workbook); err != nil {
			return err
		}
	}
	data, err := cells.WorkbookToByteData(workbook)
	if err != nil {
		return err
	}
	return sink.Write("", data)
}
