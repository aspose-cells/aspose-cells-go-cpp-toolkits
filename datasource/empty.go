package datasource

import (
	engine "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/engine"
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
	engine.LockEngine()
	defer engine.UnlockEngine()
	wb, err := engine.NewWorkbook()
	if err != nil {
		return nil, err
	}
	// Release the engine-side workbook once it has been serialized; the binding
	// would only do this from a finalizer, at some later GC.
	defer engine.CloseWorkbook(wb)

	data, err := wb.Save_SaveFormat(asposecells.SaveFormat_Xlsx)
	if err != nil {
		return nil, err
	}
	// The returned bytes are an independent copy, so they outlive the disposal.
	return BytesSource(data), nil
}
