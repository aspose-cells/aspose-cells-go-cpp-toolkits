package converter

import (
	"io"
	"os"
	"path/filepath"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/formats"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions"
)

// ConvertSpreadsheet converts a spreadsheet from the given data source into the specified output format
// and returns the resulting binary data.
//
// Parameters:
//   - source: A data source implementing the datasource.DataSource interface, which provides the input
//     spreadsheet content (e.g., from a file, in-memory buffer, HTTP URL, etc.).
//   - opt: Conversion options that define the output format and behavior, implementing the
//     saveoptions.SaveOption interface (e.g., PDFSaveOption, XLSXSaveOption, CSVSaveOption, etc.).
//
// Returns:
//   - []byte: The converted file content as a byte slice in the target format.
//   - error: An error if the conversion fails due to reasons such as unreadable source,
//     unsupported format, missing license, or failure in the underlying Aspose.Cells engine.
//
// Example:
//
// save_option := pdf.New(pdf.WithOnePagePerSheet(true))
// bytes_data, err := converter.ConvertSpreadsheet(datasource.FilePathSource("TestData/Source/BookText.xlsx"), save_option)
// if err != nil {
// println(err)
// return
// }
// os.WriteFile("TestData/Output/output2.pdf", bytes_data, 0644)
func ConvertSpreadsheet(source datasource.DataSource, opt saveoptions.SaveOption) ([]byte, error) {

	reader, errOpen := source.Open()
	if errOpen != nil {
		return nil, errOpen
	}
	data, errRead := io.ReadAll(reader)
	if errRead != nil {
		return nil, errRead
	}
	reader.Close()
	result, errApply := opt.Apply(data)

	if errApply != nil {
		return nil, errApply
	}
	return result, nil
}

// ConvertToWriter converts a spreadsheet from the given data source into the specified output format
// and writes the result directly to the provided io.Writer.
//
// Parameters:
//   - source: A data source implementing the datasource.DataSource interface, which supplies the input
//     spreadsheet content (e.g., from a file, bytes buffer, or network stream).
//   - w: An io.Writer (such as a file, bytes.Buffer, or HTTP response writer) where the converted
//     output will be written.
//   - opt: Conversion settings that determine the output format and behavior, implementing the
//     saveoptions.SaveOption interface (e.g., PDFSaveOption, XLSXSaveOption, CSVSaveOption).
//
// Returns:
//   - error: An error if conversion or writing fails. Possible causes include invalid input data,
//     unsupported output format, missing native dependencies, license restrictions in the underlying
//     Aspose.Cells engine, or write failures on the io.Writer.
//
// Example:
//
// file, err := os.Create("TestData/Output/output2.md")
// if err != nil {
// panic(err)
// }
// err = converter.ConvertToWriter(datasource.FilePathSource("TestData/Source/BookText.xlsx"), markdown.New(markdown.WithClearData(true)), file)
// if err != nil {
// return
// }
// defer file.Close()
func ConvertToWriter(source datasource.DataSource, w io.Writer, opt saveoptions.SaveOption) error {

	reader, errOpen := source.Open()
	if errOpen != nil {
		return errOpen
	}
	data, errRead := io.ReadAll(reader)
	if errRead != nil {
		return errRead
	}
	reader.Close()

	result, errApply := opt.Apply(data)
	if errApply != nil {
		return errApply
	}
	_, errWrite := w.Write(result)
	return errWrite
}

// ConvertFile converts a spreadsheet file from the input path to an output file at the specified output path.
// The output format is inferred automatically from the file extension of outputPath (e.g., ".pdf", ".xlsx", ".csv").
//
// Parameters:
//   - inputPath: Path to the source spreadsheet file (e.g., "input.xlsx"). Must be readable and in a supported format.
//   - outputPath: Path where the converted file will be written (e.g., "output.pdf"). Parent directories must exist.
//
// Returns:
//   - error: An error if the conversion fails. Possible causes include:
//   - Input file not found or unreadable,
//   - Unsupported input or output format,
//   - Invalid file content,
//   - Failure to write the output file,
//   - Underlying issues in the Aspose.Cells engine (e.g., missing native libraries or license restrictions).
//
// Example:
//
//	err := ConvertFile("data/report.xlsx", "data/report.pdf")
//	if err != nil {
//	    log.Fatalf("Conversion failed: %v", err)
//	}
func ConvertSpreadsheetToFile(inputPath string, outputPath string) error {

	source := datasource.FilePathSource(inputPath)
	ext := filepath.Ext(outputPath)
	save_option := formats.Get(ext[1:])
	data, err := ConvertSpreadsheet(source, save_option)
	if err != nil {
		return err
	}
	file, errCreate := os.Create(outputPath)
	if errCreate != nil {
		panic(errCreate)
	}
	defer file.Close()

	_, err = file.Write(data)
	if err != nil {
		panic(err)
	}

	err = file.Sync()
	return err
}
