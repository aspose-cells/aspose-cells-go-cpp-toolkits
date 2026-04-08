package splitter

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"path/filepath"
	"strings"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

func SplitSpreadsheet(source datasource.DataSource, outSaveOption saveoptions.SaveOption) ([]byte, error) {

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
	defaultStyle, _ := workbook.GetDefaultStyle()
	worksheets, _ := workbook.GetWorksheets()
	count, _ := worksheets.GetCount()
	buf := new(bytes.Buffer)
	zipWriter := zip.NewWriter(buf)
	for i := int32(0); i < count; i++ {
		worksheet, _ := worksheets.Get_Int(i)
		sheetname, _ := worksheet.GetName()
		newWorkbook, _ := asposecells.NewWorkbook()
		newWorkbook.CopyTheme(workbook)
		newDefaultStyle, _ := newWorkbook.GetDefaultStyle()
		newDefaultStyle.Copy(defaultStyle)
		newWorksheets, _ := newWorkbook.GetWorksheets()
		newWorksheet, _ := newWorksheets.Get_Int(int32(0))
		newWorksheet.SetName(sheetname)
		newWorksheet.Copy_Worksheet(worksheet)
		newFilename := sheetname + "." + outSaveOption.GetFormat()
		newWorkbook.SetFileName(newFilename)
		newData, _ := newWorkbook.SaveToStream()
		subFile, _ := zipWriter.Create(newFilename)
		outData, _ := outSaveOption.Apply(newData)
		subFile.Write(outData)
	}
	zipWriter.Close()
	return buf.Bytes(), nil
}

func SplitSpreadsheetToZipWriter(source datasource.DataSource, zipWriter zip.Writer, outSaveOption saveoptions.SaveOption) error {

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
	defaultStyle, _ := workbook.GetDefaultStyle()
	worksheets, _ := workbook.GetWorksheets()
	count, _ := worksheets.GetCount()
	for i := int32(0); i < count; i++ {
		worksheet, _ := worksheets.Get_Int(i)
		sheetname, _ := worksheet.GetName()
		newWorkbook, _ := asposecells.NewWorkbook()
		newWorkbook.CopyTheme(workbook)
		newDefaultStyle, _ := newWorkbook.GetDefaultStyle()
		newDefaultStyle.Copy(defaultStyle)
		newWorksheets, _ := newWorkbook.GetWorksheets()
		newWorksheet, _ := newWorksheets.Get_Int(int32(0))
		newWorksheet.SetName(sheetname)
		newWorksheet.Copy_Worksheet(worksheet)
		newFilename := sheetname + "." + outSaveOption.GetFormat()
		newWorkbook.SetFileName(newFilename)
		newData, _ := newWorkbook.SaveToStream()
		subFile, _ := zipWriter.Create(newFilename)
		outData, _ := outSaveOption.Apply(newData)
		subFile.Write(outData)
	}
	return nil
}
func SplitSpreadsheetToFolder(inputPath string, outputFolder string) error {

	workbook, _ := asposecells.NewWorkbook_String(inputPath)
	defaultStyle, _ := workbook.GetDefaultStyle()
	worksheets, _ := workbook.GetWorksheets()
	filename, _ := workbook.GetFileName()
	// print("filename:" + filename + ".")
	baseName := filepath.Base(filename)
	name := strings.TrimSuffix(baseName, filepath.Ext(baseName))
	ext := filepath.Ext(baseName)
	count, _ := worksheets.GetCount()
	fileFormat, _ := workbook.GetFileFormat()
	for i := int32(0); i < count; i++ {
		worksheet, _ := worksheets.Get_Int(i)
		sheetname, _ := worksheet.GetName()
		newWorkbook, _ := asposecells.NewWorkbook()
		newWorkbook.CopyTheme(workbook)
		newDefaultStyle, _ := newWorkbook.GetDefaultStyle()
		newDefaultStyle.Copy(defaultStyle)
		newFilename := fmt.Sprintf("%s_%s%s", name, sheetname, ext)
		newPath := filepath.Join(outputFolder, newFilename)
		newWorkbook.SetFileName(newFilename)
		newWorkbook.SetFileFormat(fileFormat)
		newWorksheets, _ := newWorkbook.GetWorksheets()
		newWorksheet, _ := newWorksheets.Get_Int(int32(0))
		newWorksheet.SetName(sheetname)
		newWorksheet.Copy_Worksheet(worksheet)
		newWorkbook.Save_String(newPath)
	}
	return nil
}
