package manipulator

import (
	"archive/zip"
	"bytes"
	"fmt"
	"path/filepath"
	"strings"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// renderWorksheetOutputs splits a workbook by worksheet: for each worksheet it
// builds a standalone workbook carrying the source theme and default style,
// renders it in the requested output format, and hands the (filename, data)
// pair to emit. emit is backed by a datasource.DataSink in Split.
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

// Split splits a spreadsheet by worksheet into multiple files in the output
// format described by opt, writing one (filename, data) pair per worksheet to
// sink.
//
// Parameters:
//   - source: A data source implementing the datasource.DataSource interface,
//     which provides the input spreadsheet content.
//   - opt: Output options defining the split files' format and behavior,
//     implementing the saveoptions.SaveOption interface.
//   - sink: The output destination implementing datasource.DataSink. Use
//     datasource.FolderSink for one file per worksheet, datasource.ZipSink for
//     one archive entry per worksheet, or datasource.BytesSink with a
//     datasource.ZipSink when a zip archive is required.
//
// Example:
//
//	save_option := html.New(html.WithExportImagesAsBase64(true), html.WithSaveAsSingleFile(true))
//	err := manipulator.Split(datasource.FilePathSource("examples/data/BookText.xlsx"),
//		save_option, datasource.FolderSink("out/sheets"))
func Split(source datasource.DataSource, opt saveoptions.SaveOption, sink datasource.DataSink) error {
	if opt == nil {
		return toolkiterrors.ErrSaveOptionNil
	}
	if source == nil {
		return toolkiterrors.ErrDataSourceNil
	}
	if sink == nil {
		return toolkiterrors.ErrDataSinkNil
	}
	workbook, err := cells.GetWorkbookWithDataSource(source)
	if err != nil {
		return err
	}
	return renderWorksheetOutputs(workbook, opt, sink.Write)
}

// SplitSpreadsheet splits a spreadsheet by worksheet into multiple files in the
// specified output format and returns a []byte containing all output files
// compressed into a ZIP archive.
//
// Deprecated: use Split with a datasource.ZipSink over an in-memory
// zip.Writer instead.
func SplitSpreadsheet(source datasource.DataSource, outSaveOption saveoptions.SaveOption) ([]byte, error) {
	buf := new(bytes.Buffer)
	zipWriter := zip.NewWriter(buf)
	if err := Split(source, outSaveOption, datasource.NewZipSink(zipWriter)); err != nil {
		return nil, err
	}
	if err := zipWriter.Close(); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

// SplitSpreadsheetToZipWriter splits a spreadsheet by worksheet into multiple
// files in the specified output format, writing all output files into the given
// ZIP archive.
//
// Deprecated: use Split with a datasource.ZipSink instead.
func SplitSpreadsheetToZipWriter(source datasource.DataSource, zipWriter *zip.Writer, outSaveOption saveoptions.SaveOption) error {
	return Split(source, outSaveOption, datasource.NewZipSink(zipWriter))
}

// SplitSpreadsheetToFolder splits a spreadsheet by worksheet into multiple
// files in a folder.
//
// Deprecated: use Split with a datasource.FolderSink instead. Note that Split
// names each file <sheet>.<ext>, whereas this function used the source file's
// base name as a prefix (<base>_<sheet>.<ext>).
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
