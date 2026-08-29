package transfer

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	cellsio "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/io"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

func ImportCSVDataIntoSpreadsheet(source datasource.DataSource, csvDataSource datasource.DataSource, worksheet string, beginRow int, beginColumn int, convertNumericData bool, splitter string) ([]byte, error) {
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return nil, err
	}
	worksheetCells, err := cells.GetCellsWithWorksheet(workbook, worksheet)
	if err != nil {
		return nil, err
	}
	err = worksheetCells.ImportCSV_Stream_String_Bool_Int_Int(csvDataSource.ByteData(), splitter, convertNumericData, int32(beginRow), int32(beginColumn))
	if err != nil {
		return nil, err
	}
	return cells.WorkbookToByteData(workbook)
}

func ImportXMLDataIntoSpreadsheet(source datasource.DataSource, xmlDataSource datasource.DataSource, worksheet string, beginRow int, beginColumn int) ([]byte, error) {
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return nil, err
	}
	err = workbook.ImportXml_Stream_String_Int_Int(xmlDataSource.ByteData(), worksheet, int32(beginRow), int32(beginColumn))
	if err != nil {

		return nil, err
	}
	return cells.WorkbookToByteData(workbook)
}

func ImportJsonDataIntoSpreadsheet(source datasource.DataSource, jsonDataSource datasource.DataSource, worksheet string, beginRow int, beginColumn int) ([]byte, error) {
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return nil, err
	}
	worksheetCells, err := cells.GetCellsWithWorksheet(workbook, worksheet)
	if err != nil {
		return nil, err
	}

	options, err := asposecells.NewJsonLayoutOptions()
	if err != nil {
		return nil, err
	}
	_, err = asposecells.JsonUtility_ImportData(string(jsonDataSource.ByteData()), worksheetCells, int32(beginRow), int32(beginColumn), options)
	if err != nil {
		return nil, err
	}

	return cells.WorkbookToByteData(workbook)
}

func ImportCSVFile(spreadsheet string, csvFile string, worksheet string, beginRow int, beginColumn int, convertNumericData bool, splitter string, outputPath string) error {
	data, err := ImportCSVDataIntoSpreadsheet(datasource.FilePathSource(spreadsheet), datasource.FilePathSource(csvFile), worksheet, beginRow, beginColumn, convertNumericData, splitter)
	if err != nil {
		return err
	}
	return cellsio.WriteFile(data, outputPath)
}

func ImportXMLFile(spreadsheet string, xmlFile string, worksheet string, beginRow int, beginColumn int, outputPath string) error {
	data, err := ImportXMLDataIntoSpreadsheet(datasource.FilePathSource(spreadsheet), datasource.FilePathSource(xmlFile), worksheet, beginRow, beginColumn)
	if err != nil {
		return err
	}
	return cellsio.WriteFile(data, outputPath)
}

func ImportJsonFile(spreadsheet string, jsonFile string, worksheet string, beginRow int, beginColumn int, outputPath string) error {
	data, err := ImportJsonDataIntoSpreadsheet(datasource.FilePathSource(spreadsheet), datasource.FilePathSource(jsonFile), worksheet, beginRow, beginColumn)
	if err != nil {
		return err
	}
	return cellsio.WriteFile(data, outputPath)
}
