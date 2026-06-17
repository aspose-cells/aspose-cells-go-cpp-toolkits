package transfer

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
	"io"
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

//func ExportRangeToJson(source datasource.DataSource, worksheet string , cellArea string) ([]byte, error) {
//	reader, errOpen := source.Open()
//	if errOpen != nil {
//		return nil, errOpen
//	}
//	data, errRead := io.ReadAll(reader)
//	if errRead != nil {
//		return  nil, errRead
//	}
//	reader.Close()
//	workbook, _ := asposecells.NewWorkbook_Stream(data)
//	worksheets ,err :=  workbook.GetWorksheets()
//	if err !=nil {
//		worksheet , err :=  worksheets.Get_String(worksheet)
//	}
//	return workbook.ExportXml_String_String(mapName, path)
//}
