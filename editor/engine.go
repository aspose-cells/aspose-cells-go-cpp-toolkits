package editor

import (
	"fmt"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
	"io"
)

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
func resolveSheet(wb *asposecells.Workbook, id interface{}) (*asposecells.Worksheet, error) {
	wss, _ := wb.GetWorksheets()
	switch v := id.(type) {
	case int:
		return wss.Get_Int(int32(v))
	case int32:
		return wss.Get_Int(int32(v))
	case int64:
		return wss.Get_Int(int32(v))
	case string:
		return wss.Get_String(v)
	default:
		println("default")
		return nil, fmt.Errorf("invalid sheet identifier: %v", id)
	}
	return nil, fmt.Errorf("invalid sheet identifier: %v", id)
}

func fileFormatToSaveFormat(formatType asposecells.FileFormatType) asposecells.SaveFormat {
	if formatType == asposecells.FileFormatType_Azw3 {
		return asposecells.SaveFormat_Azw3
	}
	if formatType == asposecells.FileFormatType_Bmp {
		return asposecells.SaveFormat_Bmp
	}
	if formatType == asposecells.FileFormatType_Csv {
		return asposecells.SaveFormat_Csv
	}
	if formatType == asposecells.FileFormatType_Docx {
		return asposecells.SaveFormat_Docx
	}
	if formatType == asposecells.FileFormatType_Docm {
		return asposecells.SaveFormat_Docx
	}
	if formatType == asposecells.FileFormatType_Doc {
		return asposecells.SaveFormat_Docx
	}
	if formatType == asposecells.FileFormatType_Dotm {
		return asposecells.SaveFormat_Docx
	}
	if formatType == asposecells.FileFormatType_Dif {
		return asposecells.SaveFormat_Dif
	}
	if formatType == asposecells.FileFormatType_Dbf {
		return asposecells.SaveFormat_Dbf
	}
	if formatType == asposecells.FileFormatType_Epub {
		return asposecells.SaveFormat_Epub
	}
	if formatType == asposecells.FileFormatType_Emf {
		return asposecells.SaveFormat_Emf
	}
	if formatType == asposecells.FileFormatType_Excel97To2003 {
		return asposecells.SaveFormat_Excel97To2003
	}
	if formatType == asposecells.FileFormatType_Fods {
		return asposecells.SaveFormat_Fods
	}
	if formatType == asposecells.FileFormatType_Gif {
		return asposecells.SaveFormat_Gif
	}
	if formatType == asposecells.FileFormatType_Html {
		return asposecells.SaveFormat_Html
	}
	if formatType == asposecells.FileFormatType_MHtml {
		return asposecells.SaveFormat_MHtml
	}
	if formatType == asposecells.FileFormatType_Json {
		return asposecells.SaveFormat_Json
	}
	if formatType == asposecells.FileFormatType_Jpg {
		return asposecells.SaveFormat_Jpg
	}
	if formatType == asposecells.FileFormatType_Markdown {
		return asposecells.SaveFormat_Markdown
	}
	if formatType == asposecells.FileFormatType_Numbers35 {
		return asposecells.SaveFormat_Numbers
	}
	if formatType == asposecells.FileFormatType_Numbers35 {
		return asposecells.SaveFormat_Numbers
	}
	if formatType == asposecells.FileFormatType_Ods {
		return asposecells.SaveFormat_Ods
	}
	if formatType == asposecells.FileFormatType_Ots {
		return asposecells.SaveFormat_Ots
	}
	if formatType == asposecells.FileFormatType_Png {
		return asposecells.SaveFormat_Png
	}
	if formatType == asposecells.FileFormatType_Pdf {
		return asposecells.SaveFormat_Pdf
	}
	if formatType == asposecells.FileFormatType_Ppt {
		return asposecells.SaveFormat_Pptx
	}
	if formatType == asposecells.FileFormatType_Pptx {
		return asposecells.SaveFormat_Pptx
	}
	if formatType == asposecells.FileFormatType_Ppsm {
		return asposecells.SaveFormat_Pptx
	}
	if formatType == asposecells.FileFormatType_Rtf {
		return asposecells.SaveFormat_Docx
	}
	if formatType == asposecells.FileFormatType_Rtf {
		return asposecells.SaveFormat_Docx
	}
	if formatType == asposecells.FileFormatType_Sxc {
		return asposecells.SaveFormat_Sxc
	}
	if formatType == asposecells.FileFormatType_Svg {
		return asposecells.SaveFormat_Svg
	}
	if formatType == asposecells.FileFormatType_SqlScript {
		return asposecells.SaveFormat_SqlScript
	}
	if formatType == asposecells.FileFormatType_SpreadsheetML {
		return asposecells.SaveFormat_SpreadsheetML
	}
	if formatType == asposecells.FileFormatType_Tsv {
		return asposecells.SaveFormat_Tsv
	}
	if formatType == asposecells.FileFormatType_Tiff {
		return asposecells.SaveFormat_Tiff
	}
	if formatType == asposecells.FileFormatType_Tsv {
		return asposecells.SaveFormat_Tsv
	}
	if formatType == asposecells.FileFormatType_Xlsm {
		return asposecells.SaveFormat_Xlsm
	}
	if formatType == asposecells.FileFormatType_Xlsx {
		return asposecells.SaveFormat_Xlsx
	}
	if formatType == asposecells.FileFormatType_Xlsb {
		return asposecells.SaveFormat_Xlsb
	}
	if formatType == asposecells.FileFormatType_Xlam {
		return asposecells.SaveFormat_Xlam
	}
	if formatType == asposecells.FileFormatType_Xlt {
		return asposecells.SaveFormat_Xlt
	}
	if formatType == asposecells.FileFormatType_Xltx {
		return asposecells.SaveFormat_Xltx
	}
	if formatType == asposecells.FileFormatType_Xml {
		return asposecells.SaveFormat_Xml
	}
	if formatType == asposecells.FileFormatType_Xps {
		return asposecells.SaveFormat_Xps
	}

	return asposecells.SaveFormat_Auto
}
