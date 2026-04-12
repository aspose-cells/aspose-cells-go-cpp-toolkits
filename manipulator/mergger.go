package manipulator

import (
	"io"
	"os"
	"path/filepath"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/formats"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// MergeSpreadsheets Merge multi spreadsheets into a file with the specified output format.
//
// Parameters:
//   - source: The data sources implementing the datasource.DataSource interface, which provides the input
//     spreadsheets content (e.g., from a file, in-memory buffer, etc.).
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
//	save_option = html.New(html.WithExportImagesAsBase64(true), html.WithSaveAsSingleFile(true))
//	mergedDataSource := []datasource.DataSource{datasource.FilePathSource("TestData/Source/BookText.xlsx"), datasource.FilePathSource("TestData/Source/EmployeeSalesSummary.xlsx")}
//	bytes_data, err = manipulator.MergeSpreadsheets(mergedDataSource, save_option)
//	if err != nil {
//		println(err)
//		return
//	}
//	os.WriteFile("TestData/Output/mergedOutput2.html", bytes_data, 0644)
func MergeSpreadsheets(source []datasource.DataSource, outSaveOption saveoptions.SaveOption) ([]byte, error) {
	newWorkbook, _ := asposecells.NewWorkbook()
	worksheets, _ := newWorkbook.GetWorksheets()
	worksheets.RemoveAt_Int(0)
	count := len(source)
	for i := 0; i < count; i++ {
		reader, errOpen := source[i].Open()
		if errOpen != nil {
			return nil, errOpen
		}
		data, errRead := io.ReadAll(reader)
		if errRead != nil {

			return nil, errRead
		}
		reader.Close()
		workbook, _ := asposecells.NewWorkbook_Stream(data)
		newWorkbook.Combine(workbook)
	}

	newData, _ := newWorkbook.SaveToStream()
	outData, _ := outSaveOption.Apply(newData)
	return outData, nil
}

// MergeSpreadsheetsToWriter Merge multi spreadsheets into a writer holder with the specified output format.
//
// Parameters:
//   - source: The data sources implementing the datasource.DataSource interface, which provides the input
//     spreadsheets content (e.g., from a file, in-memory buffer, etc.).
//   - w: An io.Writer (such as a file, bytes.Buffer, or HTTP response writer) where the converted
//     output will be written.
//   - outSaveOption: Split outfile options that define the output format and behavior, implementing the
//     saveoptions.SaveOption interface (e.g., PDFSaveOption, XLSXSaveOption, CSVSaveOption, etc.).
//
// Returns:
//   - error: An error if the conversion fails due to reasons such as unreadable source,
//     unsupported format, missing license, or failure in the underlying Aspose.Cells engine.
//
// Example:
//
//	save_option = html.New(html.WithExportImagesAsBase64(true), html.WithSaveAsSingleFile(true))
//	mergedDataSource := []datasource.DataSource{datasource.FilePathSource("TestData/Source/BookText.xlsx"), datasource.FilePathSource("TestData/Source/EmployeeSalesSummary.xlsx")}
//	bytes_data, err = manipulator.MergeSpreadsheets(mergedDataSource, save_option)
//	if err != nil {
//		println(err)
//		return
//	}
//	os.WriteFile("TestData/Output/mergedOutput2.html", bytes_data, 0644)
func MergeSpreadsheetsToWriter(source []datasource.DataSource, w io.Writer, outSaveOption saveoptions.SaveOption) error {
	newWorkbook, _ := asposecells.NewWorkbook()
	worksheets, _ := newWorkbook.GetWorksheets()
	worksheets.RemoveAt_Int(0)
	count := len(source)
	for i := 0; i < count; i++ {
		reader, errOpen := source[i].Open()
		if errOpen != nil {
			return errOpen
		}
		data, errRead := io.ReadAll(reader)
		if errRead != nil {
			return errRead
		}
		reader.Close()
		workbook, _ := asposecells.NewWorkbook_Stream(data)
		newWorkbook.Combine(workbook)
	}
	newData, _ := newWorkbook.SaveToStream()
	outData, errApply := outSaveOption.Apply(newData)
	if errApply != nil {
		return errApply
	}
	_, errWrite := w.Write(outData)
	return errWrite
}

// MergeSpreadsheetsToFile  Merge multi spreadsheets into a file.
//
// Parameters:
//   - inputPath: A data source implementing the datasource.DataSource interface, which provides the input
//     spreadsheet content (e.g., from a file, in-memory buffer, HTTP URL, etc.).
//   - outputPath: The output file path.
//
// Returns:
//   - error: An error if the conversion fails due to reasons such as unreadable source,
//     unsupported format, missing license, or failure in the underlying Aspose.Cells engine.
//
// Example:
//
// manipulator.MergeSpreadsheetsToFile([]string{"TestData/Source/CompanySales.xlsx", "TestData/Source/BookText.xlsx", "TestData/Source/EmployeeSalesSummary.xlsx"}, "TestData/Output/MergeBook.xlsx")
func MergeSpreadsheetsToFile(inputPaths []string, outputPath string) error {
	ext := filepath.Ext(outputPath)
	outSaveOption := formats.Get(ext[1:])

	newWorkbook, _ := asposecells.NewWorkbook()
	worksheets, _ := newWorkbook.GetWorksheets()
	worksheets.RemoveAt_Int(0)
	count := len(inputPaths)
	for i := 0; i < count; i++ {
		inputPath := inputPaths[i]
		workbook, _ := asposecells.NewWorkbook_String(inputPath)
		newWorkbook.Combine(workbook)
	}
	newData, _ := newWorkbook.SaveToStream()
	outData, errApply := outSaveOption.Apply(newData)
	if errApply != nil {
		return errApply
	}
	file, errCreate := os.Create(outputPath)
	if errCreate != nil {
		panic(errCreate)
	}
	defer file.Close()

	_, err := file.Write(outData)
	if err != nil {
		panic(err)
	}
	err = file.Sync()
	return err
}
