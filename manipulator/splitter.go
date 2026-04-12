package manipulator

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

// SplitSpreadsheet Splits a spreadsheet by worksheet into multiple files in the specified output format.Returns a []byte containing all output files compressed into a ZIP archive.
//
// Parameters:
//   - source: A data source implementing the datasource.DataSource interface, which provides the input
//     spreadsheet content (e.g., from a file, in-memory buffer, HTTP URL, etc.).
//   - outSaveOption: Split outfile options that define the output format and behavior, implementing the
//     saveoptions.SaveOption interface (e.g., PDFSaveOption, XLSXSaveOption, CSVSaveOption, etc.).
//
// Returns:
//   - []byte: The converted file content as a byte slice in the target format.
//   - error: An error if the conversion fails due to reasons such as unreadable source,
//     unsupported format, missing license, or failure in the underlying Aspose.Cells engine.
//
// Example:
//
// save_option = html.New(html.WithExportImagesAsBase64(true), html.WithSaveAsSingleFile(true))
// bytes_data, err = manipulator.SplitSpreadsheet(datasource.FilePathSource("TestData/Source/BookText.xlsx"), save_option)
// if err != nil {
// println(err)
// return
// }
// os.WriteFile("TestData/Output/output5.zip", bytes_data, 0644)

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

// SplitSpreadsheetToZipWriter Splits a spreadsheet by worksheet into multiple files in the specified output format. All output files write into a ZIP archive.
//
// Parameters:
//   - source: A data source implementing the datasource.DataSource interface, which provides the input
//     spreadsheet content (e.g., from a file, in-memory buffer, HTTP URL, etc.).
//   - outSaveOption: Split outfile options that define the output format and behavior, implementing the
//     saveoptions.SaveOption interface (e.g., PDFSaveOption, XLSXSaveOption, CSVSaveOption, etc.).
//
// Returns:
//   - error: An error if the conversion fails due to reasons such as unreadable source,
//     unsupported format, missing license, or failure in the underlying Aspose.Cells engine.
//
// Example:
//
// zipWriter := zip.NewWriter(zipFile)
// save_option = image.New(image.WithImageType("png"))
// err = manipulator.SplitSpreadsheetToZipWriter(datasource.FilePathSource("TestData/Source/BookText.xlsx"), zipWriter, save_option)
// if err != nil {
// println(err)
// return
// }
// zipWriter.Flush()
// zipWriter.Close()
// zipFile.Close()
func SplitSpreadsheetToZipWriter(source datasource.DataSource, zipWriter *zip.Writer, outSaveOption saveoptions.SaveOption) error {

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

// SplitSpreadsheetToFolder Splits a spreadsheet by worksheet into multiple files in a folder.
//
// Parameters:
//   - inputPath: A data source implementing the datasource.DataSource interface, which provides the input
//     spreadsheet content (e.g., from a file, in-memory buffer, HTTP URL, etc.).
//   - outputFolder: The output folder which store split result files.
//
// Returns:
//   - error: An error if the conversion fails due to reasons such as unreadable source,
//     unsupported format, missing license, or failure in the underlying Aspose.Cells engine.
//
// Example:
//
// manipulator.SplitSpreadsheetToFolder("TestData/Source/BookText.xlsx", "TestData/Output")
func SplitSpreadsheetToFolder(inputPath string, outputFolder string) error {
	workbook, _ := asposecells.NewWorkbook_String(inputPath)
	defaultStyle, _ := workbook.GetDefaultStyle()
	worksheets, _ := workbook.GetWorksheets()
	filename, _ := workbook.GetFileName()
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
