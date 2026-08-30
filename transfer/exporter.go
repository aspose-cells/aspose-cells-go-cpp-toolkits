// Package transfer exports spreadsheet data to structured formats and imports
// structured data back into spreadsheets.
//
// Exports cover XML, JSON, and per-worksheet JSON output; imports cover CSV,
// XML, and JSON data into a worksheet. Results are returned as bytes or written
// to a file.
package transfer

import (
	"fmt"
	"os"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	jsonsaveoptions "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/json"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

func ExportSpreadsheetToXml(source datasource.DataSource, mapName string) ([]byte, error) {
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return nil, err
	}
	return workbook.ExportXml_String(mapName)
}
func ExportRangeToJson(source datasource.DataSource, worksheet string, startCellName string, endCellName string) ([]byte, error) {
	data, err := cells.ReadSource(source)
	if err != nil {
		return nil, err
	}
	workbook, err := asposecells.NewWorkbook_Stream(data)
	if err != nil {
		return nil, err
	}
	worksheets, err := workbook.GetWorksheets()
	if err != nil {
		return nil, err
	}
	ws, err := cells.WorksheetByName(worksheets, worksheet)
	if err != nil {
		return nil, err
	}
	sheetIndex, err := ws.GetIndex()
	if err != nil {
		return nil, err
	}
	start_row_index, start_column_index, err := asposecells.CellsHelper_CellNameToIndex(startCellName)
	if err != nil {
		return nil, err
	}
	end_row_index, end_column_index, err := asposecells.CellsHelper_CellNameToIndex(endCellName)
	if err != nil {
		return nil, err
	}
	cellArea, err := asposecells.CellArea_CreateCellArea_Int_Int_Int_Int(start_row_index, start_column_index, end_row_index, end_column_index)
	if err != nil {
		return nil, err
	}
	saveoptions := jsonsaveoptions.New(jsonsaveoptions.WithSheetIndexes([]int32{sheetIndex}), jsonsaveoptions.WithExportArea(cellArea))
	return saveoptions.Apply(data)
}
func ExportWorksheetToJson(source datasource.DataSource, worksheet string) ([]byte, error) {
	data, err := cells.ReadSource(source)
	if err != nil {
		return nil, err
	}
	workbook, err := asposecells.NewWorkbook_Stream(data)
	if err != nil {
		return nil, err
	}
	sheetIndex, cellArea, err := cells.GetCellAreaWithWorksheet(workbook, worksheet)
	if err != nil {
		return nil, err
	}
	saveoptions := jsonsaveoptions.New(jsonsaveoptions.WithSheetIndexes([]int32{sheetIndex}), jsonsaveoptions.WithExportArea(cellArea))
	return saveoptions.Apply(data)
}

func ExportWorksheetToJsonFile(spreadsheet string, worksheet string, outputPath string) error {
	fileInfo, err := os.Stat(spreadsheet)
	if err != nil {
		return err
	}
	if fileInfo.IsDir() {
		return fmt.Errorf("%q is a folder, expected a file: %w", spreadsheet, toolkiterrors.ErrInputIsFolder)
	}
	data, err := ExportWorksheetToJson(datasource.FilePathSource(spreadsheet), worksheet)
	if err != nil {
		return err
	}
	return os.WriteFile(outputPath, data, 0644)
}
func ExportSpreadsheetToXmlFile(spreadsheet string, mapName string, outputPath string) error {
	fileInfo, err := os.Stat(spreadsheet)
	if err != nil {
		return err
	}
	if fileInfo.IsDir() {
		return fmt.Errorf("%q is a folder, expected a file: %w", spreadsheet, toolkiterrors.ErrInputIsFolder)
	}
	data, err := ExportSpreadsheetToXml(datasource.FilePathSource(spreadsheet), mapName)
	if err != nil {
		return err
	}
	return os.WriteFile(outputPath, data, 0644)
}
func ExportRangeToJsonFile(spreadsheet string, worksheet string, startCellName string, endCellName string, outputPath string) error {
	fileInfo, err := os.Stat(spreadsheet)
	if err != nil {
		return err
	}
	if fileInfo.IsDir() {
		return fmt.Errorf("%q is a folder, expected a file: %w", spreadsheet, toolkiterrors.ErrInputIsFolder)
	}
	data, err := ExportRangeToJson(datasource.FilePathSource(spreadsheet), worksheet, startCellName, endCellName)
	if err != nil {
		return err
	}
	return os.WriteFile(outputPath, data, 0644)
}
