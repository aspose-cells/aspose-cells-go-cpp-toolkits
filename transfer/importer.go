package transfer

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// ImportCSV imports CSV data into a worksheet of the source workbook and writes
// the resulting workbook to sink. The target sheet defaults to "Sheet1", the
// top-left cell to (0,0); use WithSheet / WithBeginCell / WithConvertNumeric /
// WithSeparator to change the defaults.
//
// Example:
//
//	err := transfer.ImportCSV(
//		datasource.FilePathSource("out/seed.xlsx"),
//		datasource.FilePathSource("examples/data/BookCsvDuplicateData.csv"),
//		datasource.FilePathSink("out/imported-csv.xlsx"),
//		transfer.WithSheet("Imported"), transfer.WithBeginCell(0, 0),
//		transfer.WithConvertNumeric(true), transfer.WithSeparator(","))
func ImportCSV(source datasource.DataSource, csvData datasource.DataSource, sink datasource.DataSink, opts ...Option) error {
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if source == nil || csvData == nil {
		return toolkiterrors.ErrDataSourceNil
	}
	if sink == nil {
		return toolkiterrors.ErrDataSinkNil
	}
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return err
	}
	ws, err := cfg.sheet.Resolve(workbook)
	if err != nil {
		return err
	}
	worksheetCells, err := ws.GetCells()
	if err != nil {
		return err
	}
	data, err := cells.ReadSource(csvData)
	if err != nil {
		return err
	}
	err = worksheetCells.ImportCSV_Stream_String_Bool_Int_Int(data, cfg.separator, cfg.convertNumeric, int32(cfg.beginRow), int32(cfg.beginColumn))
	if err != nil {
		return err
	}
	out, err := cells.WorkbookToByteData(workbook)
	if err != nil {
		return err
	}
	return sink.Write("", out)
}

// ImportJsonData imports JSON data into a worksheet of the source workbook and
// writes the resulting workbook to sink. The target sheet defaults to "Sheet1",
// the top-left cell to (0,0); use WithSheet / WithBeginCell to change them.
func ImportJsonData(source datasource.DataSource, jsonData datasource.DataSource, sink datasource.DataSink, opts ...Option) error {
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if source == nil || jsonData == nil {
		return toolkiterrors.ErrDataSourceNil
	}
	if sink == nil {
		return toolkiterrors.ErrDataSinkNil
	}
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return err
	}
	ws, err := cfg.sheet.Resolve(workbook)
	if err != nil {
		return err
	}
	worksheetCells, err := ws.GetCells()
	if err != nil {
		return err
	}
	data, err := cells.ReadSource(jsonData)
	if err != nil {
		return err
	}
	options, err := asposecells.NewJsonLayoutOptions()
	if err != nil {
		return err
	}
	_, err = asposecells.JsonUtility_ImportData(string(data), worksheetCells, int32(cfg.beginRow), int32(cfg.beginColumn), options)
	if err != nil {
		return err
	}
	out, err := cells.WorkbookToByteData(workbook)
	if err != nil {
		return err
	}
	return sink.Write("", out)
}

// ImportXMLData imports XML data into a worksheet of the source workbook and
// writes the resulting workbook to sink. The target sheet defaults to "Sheet1",
// the top-left cell to (0,0); use WithSheet / WithBeginCell to change them.
func ImportXMLData(source datasource.DataSource, xmlData datasource.DataSource, sink datasource.DataSink, opts ...Option) error {
	cfg := defaultOptions()
	applyOptions(cfg, opts)
	if source == nil || xmlData == nil {
		return toolkiterrors.ErrDataSourceNil
	}
	if sink == nil {
		return toolkiterrors.ErrDataSinkNil
	}
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return err
	}
	ws, err := cfg.sheet.Resolve(workbook)
	if err != nil {
		return err
	}
	targetName, err := ws.GetName()
	if err != nil {
		return err
	}
	data, err := cells.ReadSource(xmlData)
	if err != nil {
		return err
	}
	err = workbook.ImportXml_Stream_String_Int_Int(data, targetName, int32(cfg.beginRow), int32(cfg.beginColumn))
	if err != nil {
		return err
	}
	out, err := cells.WorkbookToByteData(workbook)
	if err != nil {
		return err
	}
	return sink.Write("", out)
}

// ImportCSVDataIntoSpreadsheet imports CSV data into a worksheet and returns
// the resulting workbook bytes.
//
// Deprecated: use ImportCSV with a datasource.BytesSink instead.
func ImportCSVDataIntoSpreadsheet(source datasource.DataSource, csvDataSource datasource.DataSource, worksheet string, beginRow int, beginColumn int, convertNumericData bool, splitter string) ([]byte, error) {
	var out datasource.BytesSink
	err := ImportCSV(source, csvDataSource, &out,
		WithSheet(worksheet), WithBeginCell(beginRow, beginColumn),
		WithConvertNumeric(convertNumericData), WithSeparator(splitter))
	if err != nil {
		return nil, err
	}
	return out.Bytes(), nil
}

// ImportCSVFile imports CSV data into a worksheet straight from file to file.
//
// Deprecated: use ImportCSV with a datasource.FilePathSource and
// datasource.FilePathSink instead.
func ImportCSVFile(spreadsheet string, csvFile string, worksheet string, beginRow int, beginColumn int, convertNumericData bool, splitter string, outputPath string) error {
	return ImportCSV(
		datasource.FilePathSource(spreadsheet),
		datasource.FilePathSource(csvFile),
		datasource.FilePathSink(outputPath),
		WithSheet(worksheet), WithBeginCell(beginRow, beginColumn),
		WithConvertNumeric(convertNumericData), WithSeparator(splitter),
	)
}

// ImportJsonDataIntoSpreadsheet imports JSON data into a worksheet and returns
// the resulting workbook bytes.
//
// Deprecated: use ImportJsonData with a datasource.BytesSink instead.
func ImportJsonDataIntoSpreadsheet(source datasource.DataSource, jsonDataSource datasource.DataSource, worksheet string, beginRow int, beginColumn int) ([]byte, error) {
	var out datasource.BytesSink
	err := ImportJsonData(source, jsonDataSource, &out, WithSheet(worksheet), WithBeginCell(beginRow, beginColumn))
	if err != nil {
		return nil, err
	}
	return out.Bytes(), nil
}

// ImportJsonFile imports JSON data into a worksheet straight from file to file.
//
// Deprecated: use ImportJsonData with a datasource.FilePathSource and
// datasource.FilePathSink instead.
func ImportJsonFile(spreadsheet string, jsonFile string, worksheet string, beginRow int, beginColumn int, outputPath string) error {
	return ImportJsonData(
		datasource.FilePathSource(spreadsheet),
		datasource.FilePathSource(jsonFile),
		datasource.FilePathSink(outputPath),
		WithSheet(worksheet), WithBeginCell(beginRow, beginColumn),
	)
}

// ImportXMLDataIntoSpreadsheet imports XML data into a worksheet and returns
// the resulting workbook bytes.
//
// Deprecated: use ImportXMLData with a datasource.BytesSink instead.
func ImportXMLDataIntoSpreadsheet(source datasource.DataSource, xmlDataSource datasource.DataSource, worksheet string, beginRow int, beginColumn int) ([]byte, error) {
	var out datasource.BytesSink
	err := ImportXMLData(source, xmlDataSource, &out, WithSheet(worksheet), WithBeginCell(beginRow, beginColumn))
	if err != nil {
		return nil, err
	}
	return out.Bytes(), nil
}

// ImportXMLFile imports XML data into a worksheet straight from file to file.
//
// Deprecated: use ImportXMLData with a datasource.FilePathSource and
// datasource.FilePathSink instead.
func ImportXMLFile(spreadsheet string, xmlFile string, worksheet string, beginRow int, beginColumn int, outputPath string) error {
	return ImportXMLData(
		datasource.FilePathSource(spreadsheet),
		datasource.FilePathSource(xmlFile),
		datasource.FilePathSink(outputPath),
		WithSheet(worksheet), WithBeginCell(beginRow, beginColumn),
	)
}
