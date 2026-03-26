package converter

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/dbf"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/docx"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/ebook"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/html"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/image"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/json"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/markdown"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/ods"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/ooxml"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/pcl"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/pdf"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/pptx"
	sqlscript "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/sqlScript"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/txt"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/xls"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/xlsb"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/xml"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/xps"
	"io"
	"os"
	"path/filepath"
	"strings"
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
//   - opt: Conversion settings that determine the output format and behavior, implementing the
//     saveoptions.SaveOption interface (e.g., PDFSaveOption, XLSXSaveOption, CSVSaveOption).
//   - w: An io.Writer (such as a file, bytes.Buffer, or HTTP response writer) where the converted
//     output will be written.
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
func ConvertToWriter(source datasource.DataSource, opt saveoptions.SaveOption, w io.Writer) error {

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
func ConvertFile(inputPath string, outputPath string) error {

	source := datasource.FilePathSource(inputPath)
	ext := filepath.Ext(outputPath)
	save_option := FromFileExt(ext[1:])
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

type SaveOptionFactory func() saveoptions.SaveOption

var formatRegistry = make(map[string]SaveOptionFactory)

func register(ext string, factory SaveOptionFactory) {
	formatRegistry[strings.ToLower(ext)] = factory
}

func FromFileExt(ext string) saveoptions.SaveOption {
	if factory, ok := formatRegistry[ext]; ok {
		return factory()
	}
	return nil
}

func init() {
	register("dbf", func() saveoptions.SaveOption {
		return dbf.New()
	})
	register("docx", func() saveoptions.SaveOption {
		return docx.New()
	})
	register("epub", func() saveoptions.SaveOption {
		return ebook.New()
	})
	register("html", func() saveoptions.SaveOption {
		return html.New()
	})
	register("png", func() saveoptions.SaveOption {
		return image.New(image.WithImageType("png"))
	})
	register("jpg", func() saveoptions.SaveOption {
		return image.New(image.WithImageType("jpg"))
	})
	register("svg", func() saveoptions.SaveOption {
		return image.New(image.WithImageType("svg"))
	})
	register("bmp", func() saveoptions.SaveOption {
		return image.New(image.WithImageType("bmp"))
	})
	register("tif", func() saveoptions.SaveOption {
		return image.New(image.WithImageType("tif"))
	})
	register("tiff", func() saveoptions.SaveOption {
		return image.New(image.WithImageType("tiff"))
	})
	register("json", func() saveoptions.SaveOption {
		return json.New()
	})
	register("md", func() saveoptions.SaveOption {
		return markdown.New()
	})
	register("ods", func() saveoptions.SaveOption {
		return ods.New()
	})
	register("xlsx", func() saveoptions.SaveOption {
		return ooxml.New()
	})
	register("xlsm", func() saveoptions.SaveOption {
		return ooxml.New()
	})
	register("pcl", func() saveoptions.SaveOption {
		return pcl.New()
	})
	register("pdf", func() saveoptions.SaveOption {
		return pdf.New()
	})
	register("pptx", func() saveoptions.SaveOption {
		return pptx.New()
	})
	register("sql", func() saveoptions.SaveOption {
		return sqlscript.New()
	})
	register("csv", func() saveoptions.SaveOption {
		return txt.New()
	})
	register("txt", func() saveoptions.SaveOption {
		return txt.New()
	})
	register("xls", func() saveoptions.SaveOption {
		return xls.New()
	})
	register("xlsb", func() saveoptions.SaveOption {
		return xlsb.New()
	})
	register("xml", func() saveoptions.SaveOption {
		return xml.New()
	})
	register("xps", func() saveoptions.SaveOption {
		return xps.New()
	})
}
