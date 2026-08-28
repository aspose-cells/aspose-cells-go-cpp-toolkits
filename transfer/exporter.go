package transfer

import (
	"fmt"
	. "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
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
			println("sheetIndex:", sheetIndex, startCellName, endCellName)
			cellArea, err := asposecells.CellArea_CreateCellArea_String_String(startCellName, endCellName)
			if err == nil {
				println(cellArea.ToString())
			}
			name, err := cellArea.ToString()
			if err == nil {
				println(name)
			}

			saveoptions := jsonsaveoptions.New(jsonsaveoptions.WithSheetIndexes([]int32{sheetIndex}), jsonsaveoptions.WithExportArea(cellArea))
			return saveoptions.Apply(data)
		}
	}
	return nil, err
}
func ExportWorksheetToJson(source datasource.DataSource, worksheet string) ([]byte, error) {

	workbook, err := GetWorkbookWithDataSource(source)
	if err != nil {
		return nil, err
	}
	sheetIndex, cellArea, err := GetCellAreaWithWorksheet(workbook, worksheet)
	if err != nil {
		return nil, err
	}
	saveoptions := jsonsaveoptions.New(jsonsaveoptions.WithSheetIndexes([]int32{sheetIndex}), jsonsaveoptions.WithExportArea(cellArea))
	return saveoptions.Apply(source.ByteData())
}

func ExportWorksheetToJsonFile(spreadsheet string, worksheet string, outputPath string) error {
	fileInfo, err := os.Stat(spreadsheet)
	if err != nil {
		return err
	}
	if fileInfo.IsDir() {
		return fmt.Errorf("The %s is folder.", spreadsheet)
	}
	data, err := ExportWorksheetToJson(datasource.FilePathSource(spreadsheet), worksheet)
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
func ExportSpreadsheetToXmlFile(spreadsheet string, mapName string, outputPath string) error {
	fileInfo, err := os.Stat(spreadsheet)
	if err != nil {
		return err
	}
	if fileInfo.IsDir() {
		return fmt.Errorf("The %s is folder.", spreadsheet)
	}
	data, err := ExportSpreadsheetToXml(datasource.FilePathSource(spreadsheet), mapName)
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
func ExportRangeToJsonFile(spreadsheet string, worksheet string, startCellName string, endCellName string, outputPath string) error {
	fileInfo, err := os.Stat(spreadsheet)
	if err != nil {
		return err
	}
	if fileInfo.IsDir() {
		return fmt.Errorf("The %s is folder.", spreadsheet)
	}
	data, err := ExportRangeToJson(datasource.FilePathSource(spreadsheet), worksheet, startCellName, endCellName)
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
