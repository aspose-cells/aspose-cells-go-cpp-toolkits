// Package manipulator merges and splits spreadsheets.
//
// MergeSpreadsheets combines multiple input sources into one workbook; the
// Split* functions export every worksheet of a workbook as a standalone file.
// Output is returned as bytes, streamed to a zip writer, or written to a folder.
package manipulator

import (
	"fmt"
	"io"
	"os"
	"path/filepath"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
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
	if outSaveOption == nil {
		return nil, toolkiterrors.ErrSaveOptionNil
	}
	newWorkbook, err := asposecells.NewWorkbook()
	if err != nil {
		return nil, err
	}
	worksheets, err := newWorkbook.GetWorksheets()
	if err != nil {
		return nil, err
	}
	err = worksheets.RemoveAt_Int(0)
	if err != nil {
		return nil, err
	}
	count := len(source)
	for i := 0; i < count; i++ {
		workbook, err := cells.GetWorkbookWithDataSource(source[i])
		if err != nil {
			return nil, err
		}
		err = newWorkbook.Combine(workbook)
		if err != nil {
			return nil, err
		}
	}

	newData, err := newWorkbook.SaveToStream()
	if err != nil {
		return nil, err
	}
	outData, err := outSaveOption.Apply(newData)
	if err != nil {
		return nil, err
	}
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
	if outSaveOption == nil {
		return toolkiterrors.ErrSaveOptionNil
	}
	newWorkbook, err := asposecells.NewWorkbook()
	if err != nil {
		return err
	}
	worksheets, err := newWorkbook.GetWorksheets()
	if err != nil {
		return err
	}
	err = worksheets.RemoveAt_Int(0)
	if err != nil {
		return err
	}
	count := len(source)
	for i := 0; i < count; i++ {
		workbook, err := cells.GetWorkbookWithDataSource(source[i])
		if err != nil {
			return err
		}
		err = newWorkbook.Combine(workbook)
		if err != nil {
			return err
		}
	}
	newData, err := newWorkbook.SaveToStream()
	if err != nil {
		return err
	}
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
	if len(ext) <= 1 {
		return fmt.Errorf("invalid output path %q: missing file extension: %w", outputPath, toolkiterrors.ErrInvalidOutputPath)
	}
	outSaveOption := formats.Get(ext[1:])
	if outSaveOption == nil {
		return fmt.Errorf("unsupported output format %q: %w", ext[1:], toolkiterrors.ErrUnsupportedFormat)
	}

	newWorkbook, err := asposecells.NewWorkbook()
	if err != nil {
		return err
	}
	worksheets, err := newWorkbook.GetWorksheets()
	if err != nil {
		return err
	}
	err = worksheets.RemoveAt_Int(0)
	if err != nil {
		return err
	}
	count := len(inputPaths)
	for i := 0; i < count; i++ {
		workbook, err := asposecells.NewWorkbook_String(inputPaths[i])
		if err != nil {
			return err
		}
		err = newWorkbook.Combine(workbook)
		if err != nil {
			return err
		}
	}
	newData, err := newWorkbook.SaveToStream()
	if err != nil {
		return err
	}
	outData, err := outSaveOption.Apply(newData)
	if err != nil {
		return err
	}
	return os.WriteFile(outputPath, outData, 0644)
}
