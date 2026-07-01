package transfer

import (
	"io"
	"os"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	jsonsaveoptions "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/json"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

func ExportSpreadsheetToXml(source datasource.DataSource, mapName string) ([]byte, error) {
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
	return workbook.ExportXml_String(mapName)
}

func ExportSpreadsheetToXmlFile(source datasource.DataSource, mapName string, path string) error {
	reader, errOpen := source.Open()
	if errOpen != nil {
		return errOpen
	}
	data, errRead := io.ReadAll(reader)
	if errRead != nil {
		return errRead
	}
	reader.Close()
	workbook, _ := asposecells.NewWorkbook_Stream(data)
	return workbook.ExportXml_String_String(mapName, path)
}

func ExportWorksheetToJson(source datasource.DataSource, worksheet string) ([]byte, error) {
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
	worksheets, err := workbook.GetWorksheets()
	if err == nil {
		worksheet, err := worksheets.Get_String(worksheet)
		if err == nil {
			cells, _ := worksheet.GetCells()
			dataColumnIndex, _ := cells.GetMaxDataColumn()
			dataRowIndex, _ := cells.GetMaxDataRow()
			endCellName, _ := asposecells.CellsHelper_CellIndexToName(dataRowIndex, dataColumnIndex)
			sheetIndex, _ := worksheet.GetIndex()
			cellArea, _ := asposecells.CellArea_CreateCellArea_String_String("A1", endCellName)
			saveoptions := jsonsaveoptions.New(jsonsaveoptions.WithSheetIndexes([]int32{sheetIndex}), jsonsaveoptions.WithExportArea(cellArea))
			return saveoptions.Apply(data)
		}
	}
	return nil, err
}
func ExportWorksheetToJsonFile(source datasource.DataSource, worksheet string, outputPath string) error {
	data, err := ExportWorksheetToJson(source, worksheet)
	if err == nil {
		file, errCreate := os.Create(outputPath)
		if errCreate != nil {
			panic(errCreate)
		}
		defer file.Close()

		_, err = file.Write(data)
		if err != nil {
			panic(err)
		}
	}
	return err
}

func ExportRangeToJson(source datasource.DataSource, worksheet string, startCellName string, endCellName string) ([]byte, error) {
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
	worksheets, err := workbook.GetWorksheets()
	if err == nil {
		worksheet, err := worksheets.Get_String(worksheet)
		if err == nil {
			sheetIndex, _ := worksheet.GetIndex()
			cellArea, _ := asposecells.CellArea_CreateCellArea_String_String(startCellName, endCellName)
			saveoptions := jsonsaveoptions.New(jsonsaveoptions.WithSheetIndexes([]int32{sheetIndex}), jsonsaveoptions.WithExportArea(cellArea))
			return saveoptions.Apply(data)
		}
	}
	return nil, err
}
func ExportRangeToJsonFile(source datasource.DataSource, worksheet string, startCellName string, endCellName string, outputPath string) error {
	data, err := ExportRangeToJson(source, worksheet, startCellName, endCellName)
	if err == nil {
		file, errCreate := os.Create(outputPath)
		if errCreate != nil {
			panic(errCreate)
		}
		defer file.Close()

		_, err = file.Write(data)
		if err != nil {
			panic(err)
		}
	}
	return err
}
