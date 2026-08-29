package manipulator

import (
	"archive/zip"
	"bytes"
	"fmt"
	"path/filepath"
	"strings"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// renderWorksheetOutputs splits a workbook by worksheet: for each worksheet it builds a
// standalone workbook carrying the source theme and default style, renders it in the
// requested output format, and hands the (filename, data) pair to emit.
//
// This is the shared implementation behind SplitSpreadsheet and
// SplitSpreadsheetToZipWriter, so both split paths behave identically.
func renderWorksheetOutputs(workbook *asposecells.Workbook, outSaveOption saveoptions.SaveOption, emit func(filename string, data []byte) error) error {
	defaultStyle, err := workbook.GetDefaultStyle()
	if err != nil {
		return err
	}
	worksheets, err := workbook.GetWorksheets()
	if err != nil {
		return err
	}
	count, err := worksheets.GetCount()
	if err != nil {
		return err
	}
	for i := int32(0); i < count; i++ {
		worksheet, err := worksheets.Get_Int(i)
		if err != nil {
			return err
		}
		sheetname, err := worksheet.GetName()
		if err != nil {
			return err
		}
		newWorkbook, err := asposecells.NewWorkbook()
		if err != nil {
			return err
		}
		err = newWorkbook.CopyTheme(workbook)
		if err != nil {
			return err
		}
		newDefaultStyle, err := newWorkbook.GetDefaultStyle()
		if err != nil {
			return err
		}
		err = newDefaultStyle.Copy(defaultStyle)
		if err != nil {
			return err
		}
		newWorksheets, err := newWorkbook.GetWorksheets()
		if err != nil {
			return err
		}
		newWorksheet, err := newWorksheets.Get_Int(int32(0))
		if err != nil {
			return err
		}
		err = newWorksheet.SetName(sheetname)
		if err != nil {
			return err
		}
		err = newWorksheet.Copy_Worksheet(worksheet)
		if err != nil {
			return err
		}
		newFilename := sheetname + "." + outSaveOption.GetFormat()
		err = newWorkbook.SetFileName(newFilename)
		if err != nil {
			return err
		}
		newData, err := newWorkbook.SaveToStream()
		if err != nil {
			return err
		}
		outData, err := outSaveOption.Apply(newData)
		if err != nil {
			return err
		}
		if err := emit(newFilename, outData); err != nil {
			return err
		}
	}
	return nil
}

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
	if outSaveOption == nil {
		return nil, fmt.Errorf("save option is nil")
	}
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return nil, err
	}
	buf := new(bytes.Buffer)
	zipWriter := zip.NewWriter(buf)
	if err := renderWorksheetOutputs(workbook, outSaveOption, func(filename string, data []byte) error {
		subFile, err := zipWriter.Create(filename)
		if err != nil {
			return err
		}
		_, err = subFile.Write(data)
		return err
	}); err != nil {
		return nil, err
	}
	if err := zipWriter.Close(); err != nil {
		return nil, err
	}
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
	if outSaveOption == nil {
		return fmt.Errorf("save option is nil")
	}
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return err
	}
	return renderWorksheetOutputs(workbook, outSaveOption, func(filename string, data []byte) error {
		subFile, err := zipWriter.Create(filename)
		if err != nil {
			return err
		}
		_, err = subFile.Write(data)
		return err
	})
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
	workbook, err := asposecells.NewWorkbook_String(inputPath)
	if err != nil {
		return err
	}
	defaultStyle, err := workbook.GetDefaultStyle()
	if err != nil {
		return err
	}
	worksheets, err := workbook.GetWorksheets()
	if err != nil {
		return err
	}
	filename, err := workbook.GetFileName()
	if err != nil {
		return err
	}
	baseName := filepath.Base(filename)
	name := strings.TrimSuffix(baseName, filepath.Ext(baseName))
	ext := filepath.Ext(baseName)
	count, err := worksheets.GetCount()
	if err != nil {
		return err
	}
	fileFormat, err := workbook.GetFileFormat()
	if err != nil {
		return err
	}
	for i := int32(0); i < count; i++ {
		worksheet, err := worksheets.Get_Int(i)
		if err != nil {
			return err
		}
		sheetname, err := worksheet.GetName()
		if err != nil {
			return err
		}
		newWorkbook, err := asposecells.NewWorkbook()
		if err != nil {
			return err
		}
		err = newWorkbook.CopyTheme(workbook)
		if err != nil {
			return err
		}
		newDefaultStyle, err := newWorkbook.GetDefaultStyle()
		if err != nil {
			return err
		}
		err = newDefaultStyle.Copy(defaultStyle)
		if err != nil {
			return err
		}
		newFilename := fmt.Sprintf("%s_%s%s", name, sheetname, ext)
		newPath := filepath.Join(outputFolder, newFilename)
		err = newWorkbook.SetFileName(newFilename)
		if err != nil {
			return err
		}
		err = newWorkbook.SetFileFormat(fileFormat)
		if err != nil {
			return err
		}
		newWorksheets, err := newWorkbook.GetWorksheets()
		if err != nil {
			return err
		}
		newWorksheet, err := newWorksheets.Get_Int(int32(0))
		if err != nil {
			return err
		}
		err = newWorksheet.SetName(sheetname)
		if err != nil {
			return err
		}
		err = newWorksheet.Copy_Worksheet(worksheet)
		if err != nil {
			return err
		}
		err = newWorkbook.Save_String(newPath)
		if err != nil {
			return err
		}
	}
	return nil
}
