package converter

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions/dbf"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions/docx"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions/ebook"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions/html"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions/image"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions/json"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions/markdown"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions/ods"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions/ooxml"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions/pcl"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions/pdf"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions/pptx"
	sqlscript "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions/sqlScript"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions/txt"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions/xls"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions/xlsb"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions/xml"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions/xps"
	"io"
	"os"
	"path/filepath"
	"strings"
)

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
