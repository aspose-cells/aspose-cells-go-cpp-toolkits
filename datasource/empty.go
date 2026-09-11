package datasource

import (
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// NewEmptyWorkbook creates a brand-new, blank XLSX workbook entirely in
// memory and returns it as a DataSource. It is the toolkit's sanctioned way
// to "start from nothing": the result can be passed to any entry point that
// accepts a DataSource — for example editor.EditSpreadsheet to author a table
// with the DSL, or transfer.ImportCSV to import data into a fresh file.
//
// The workbook contains a single default worksheet at index 0 and is created
// without touching the disk. Unlike loading an existing file, building it
// never consumes the engine's evaluation-mode load budget.
func NewEmptyWorkbook() (BytesSource, error) {
	wb, err := asposecells.NewWorkbook()
	if err != nil {
		return nil, err
	}
	data, err := wb.Save_SaveFormat(asposecells.SaveFormat_Xlsx)
	if err != nil {
		return nil, err
	}
	return BytesSource(data), nil
}
