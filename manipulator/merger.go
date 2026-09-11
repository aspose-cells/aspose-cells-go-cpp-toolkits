// Package manipulator merges and splits spreadsheets.
//
// Merge combines multiple input sources into one workbook; Split exports every
// worksheet of a workbook as a standalone file. Both write their result to a
// datasource.DataSink, so the caller picks the output shape (file, writer,
// bytes, folder, or zip archive) by choosing the sink.
package manipulator

import (
	"fmt"
	"io"
	"path/filepath"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// Merge combines multiple spreadsheets into a single workbook in the format
// described by opt and writes the result to sink.
//
// Parameters:
//   - source: The data sources implementing the datasource.DataSource
//     interface, which provide the input spreadsheets content (e.g., from
//     files, in-memory buffers, etc.).
//   - opt: Output options defining the merged workbook's format and behavior,
//     implementing the saveoptions.SaveOption interface.
//   - sink: The output destination implementing datasource.DataSink.
//
// Example:
//
//	save_option := html.New(html.WithExportImagesAsBase64(true), html.WithSaveAsSingleFile(true))
//	mergedDataSource := []datasource.DataSource{
//		datasource.FilePathSource("examples/data/BookText.xlsx"),
//		datasource.FilePathSource("examples/data/EmployeeSalesSummary.xlsx"),
//	}
//	err := manipulator.Merge(mergedDataSource, save_option,
//		datasource.FilePathSink("out/mergedOutput2.html"))
func Merge(sources []datasource.DataSource, opt saveoptions.SaveOption, sink datasource.DataSink) error {
	if opt == nil {
		return toolkiterrors.ErrSaveOptionNil
	}
	if sink == nil {
		return toolkiterrors.ErrDataSinkNil
	}
	if len(sources) == 0 {
		return toolkiterrors.ErrNoSources
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
	for i := 0; i < len(sources); i++ {
		if sources[i] == nil {
			return fmt.Errorf("source %d is nil: %w", i, toolkiterrors.ErrDataSourceNil)
		}
		workbook, err := cells.GetWorkbookWithDataSource(sources[i])
		if err != nil {
			return fmt.Errorf("source %d: %w", i, err)
		}
		err = newWorkbook.Combine(workbook)
		if err != nil {
			return fmt.Errorf("combine source %d: %w", i, err)
		}
	}

	newData, err := newWorkbook.SaveToStream()
	if err != nil {
		return err
	}
	outData, err := opt.Apply(newData)
	if err != nil {
		return err
	}
	return sink.Write("", outData)
}

// MergeSpreadsheets merges multiple spreadsheets into a file with the specified
// output format and returns the resulting binary data.
//
// Deprecated: use Merge with a datasource.BytesSink instead.
func MergeSpreadsheets(sources []datasource.DataSource, opt saveoptions.SaveOption) ([]byte, error) {
	var out datasource.BytesSink
	if err := Merge(sources, opt, &out); err != nil {
		return nil, err
	}
	return out.Bytes(), nil
}

// MergeSpreadsheetsToWriter merges multiple spreadsheets into a writer with the
// specified output format.
//
// Deprecated: use Merge with a datasource.WriterSink instead.
func MergeSpreadsheetsToWriter(sources []datasource.DataSource, w io.Writer, opt saveoptions.SaveOption) error {
	return Merge(sources, opt, datasource.NewWriterSink(w))
}

// MergeSpreadsheetsToFile merges multiple spreadsheet files into a single
// output file. The output format is inferred from the output file's extension.
//
// Deprecated: use Merge with datasource.FilePathSource inputs and a
// datasource.FilePathSink instead.
func MergeSpreadsheetsToFile(inputPaths []string, outputPath string) error {
	ext := filepath.Ext(outputPath)
	if len(ext) <= 1 {
		return fmt.Errorf("invalid output path %q: missing file extension: %w", outputPath, toolkiterrors.ErrInvalidOutputPath)
	}
	opt := formats.Get(ext[1:])
	if opt == nil {
		return fmt.Errorf("unsupported output format %q: %w", ext[1:], toolkiterrors.ErrUnsupportedFormat)
	}
	sources := make([]datasource.DataSource, 0, len(inputPaths))
	for _, p := range inputPaths {
		sources = append(sources, datasource.FilePathSource(p))
	}
	return Merge(sources, opt, datasource.FilePathSink(outputPath))
}
