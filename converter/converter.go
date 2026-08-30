// Package converter converts spreadsheets between formats.
//
// It drives the Aspose.Cells engine over a datasource.DataSource input and
// writes the result to a datasource.DataSink, so the caller chooses the output
// shape (file, writer, or in-memory bytes) by picking the sink rather than a
// separate entry point. Format selection is delegated to a
// saveoptions.SaveOption, so every supported target format shares the same
// entry point.
package converter

import (
	"fmt"
	"io"
	"path/filepath"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
)

// Convert converts a spreadsheet from the given data source into the format
// described by opt and writes the result to sink.
//
// Parameters:
//   - source: A data source implementing the datasource.DataSource interface,
//     which provides the input spreadsheet content (e.g., from a file,
//     in-memory buffer, HTTP URL, etc.).
//   - opt: Conversion options that define the output format and behavior,
//     implementing the saveoptions.SaveOption interface (e.g., PDFSaveOption,
//     XLSXSaveOption, CSVSaveOption, etc.).
//   - sink: The output destination implementing datasource.DataSink. Use
//     datasource.FilePathSink for a file, datasource.WriterSink for an
//     io.Writer, or datasource.BytesSink to capture the result as bytes.
//
// Example:
//
//	save_option := pdf.New(pdf.WithOnePagePerSheet(true))
//	err := converter.Convert(datasource.FilePathSource("examples/data/BookText.xlsx"),
//		save_option, datasource.FilePathSink("out/output2.pdf"))
func Convert(source datasource.DataSource, opt saveoptions.SaveOption, sink datasource.DataSink) error {
	if opt == nil {
		return toolkiterrors.ErrSaveOptionNil
	}
	if source == nil {
		return toolkiterrors.ErrDataSourceNil
	}
	if sink == nil {
		return toolkiterrors.ErrDataSinkNil
	}
	data, errRead := cells.ReadSource(source)
	if errRead != nil {
		return errRead
	}
	result, errApply := opt.Apply(data)
	if errApply != nil {
		return errApply
	}
	return sink.Write("", result)
}

// ConvertSpreadsheet converts a spreadsheet from the given data source into the
// specified output format and returns the resulting binary data.
//
// Deprecated: use Convert with a datasource.BytesSink instead:
//
//	var out datasource.BytesSink
//	if err := Convert(source, opt, &out); err != nil {
//		return nil, err
//	}
//	bytes_data := out.Bytes()
func ConvertSpreadsheet(source datasource.DataSource, opt saveoptions.SaveOption) ([]byte, error) {
	var out datasource.BytesSink
	if err := Convert(source, opt, &out); err != nil {
		return nil, err
	}
	return out.Bytes(), nil
}

// ConvertToWriter converts a spreadsheet from the given data source into the
// specified output format and writes the result directly to the provided
// io.Writer.
//
// Deprecated: use Convert with a datasource.WriterSink instead:
//
//	err := Convert(source, opt, datasource.NewWriterSink(w))
func ConvertToWriter(source datasource.DataSource, w io.Writer, opt saveoptions.SaveOption) error {
	return Convert(source, opt, datasource.NewWriterSink(w))
}

// ConvertSpreadsheetToFile converts a spreadsheet file from the input path to
// an output file at the specified output path. The output format is inferred
// automatically from the file extension of outputPath.
//
// Deprecated: use Convert with a datasource.FilePathSource and
// datasource.FilePathSink instead:
//
//	opt := formats.Get(filepath.Ext(outputPath)[1:])
//	err := Convert(datasource.FilePathSource(inputPath), opt,
//		datasource.FilePathSink(outputPath))
func ConvertSpreadsheetToFile(inputPath string, outputPath string) error {
	source := datasource.FilePathSource(inputPath)
	ext := filepath.Ext(outputPath)
	if len(ext) <= 1 {
		return fmt.Errorf("invalid output path %q: missing file extension: %w", outputPath, toolkiterrors.ErrInvalidOutputPath)
	}
	opt := formats.Get(ext[1:])
	if opt == nil {
		return fmt.Errorf("unsupported output format %q: %w", ext[1:], toolkiterrors.ErrUnsupportedFormat)
	}
	return Convert(source, opt, datasource.FilePathSink(outputPath))
}
