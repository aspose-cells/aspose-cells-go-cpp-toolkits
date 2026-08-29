package formats

import (
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
	"strings"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
)

var registry = make(map[string]func() saveoptions.SaveOption)

func Register(ext string, factory func() saveoptions.SaveOption) {
	registry[strings.ToLower(ext)] = factory
}

func Get(ext string) saveoptions.SaveOption {
	if factory, ok := registry[strings.ToLower(ext)]; ok {
		return factory()
	}
	return nil
}

func List() []string {
	exts := make([]string, 0, len(registry))
	for ext := range registry {
		exts = append(exts, ext)
	}
	return exts
}

func FileFormatToSaveFormat(formatType asposecells.FileFormatType) asposecells.SaveFormat {
	switch formatType {
	case asposecells.FileFormatType_Azw3:
		return asposecells.SaveFormat_Azw3
	case asposecells.FileFormatType_Bmp:
		return asposecells.SaveFormat_Bmp
	case asposecells.FileFormatType_Csv:
		return asposecells.SaveFormat_Csv
	case asposecells.FileFormatType_Docx,
		asposecells.FileFormatType_Docm,
		asposecells.FileFormatType_Doc,
		asposecells.FileFormatType_Dotm,
		asposecells.FileFormatType_Rtf:
		return asposecells.SaveFormat_Docx
	case asposecells.FileFormatType_Dif:
		return asposecells.SaveFormat_Dif
	case asposecells.FileFormatType_Dbf:
		return asposecells.SaveFormat_Dbf
	case asposecells.FileFormatType_Epub:
		return asposecells.SaveFormat_Epub
	case asposecells.FileFormatType_Emf:
		return asposecells.SaveFormat_Emf
	case asposecells.FileFormatType_Excel97To2003:
		return asposecells.SaveFormat_Excel97To2003
	case asposecells.FileFormatType_Fods:
		return asposecells.SaveFormat_Fods
	case asposecells.FileFormatType_Gif:
		return asposecells.SaveFormat_Gif
	case asposecells.FileFormatType_Html:
		return asposecells.SaveFormat_Html
	case asposecells.FileFormatType_MHtml:
		return asposecells.SaveFormat_MHtml
	case asposecells.FileFormatType_Json:
		return asposecells.SaveFormat_Json
	case asposecells.FileFormatType_Jpg:
		return asposecells.SaveFormat_Jpg
	case asposecells.FileFormatType_Markdown:
		return asposecells.SaveFormat_Markdown
	case asposecells.FileFormatType_Numbers35:
		return asposecells.SaveFormat_Numbers
	case asposecells.FileFormatType_Ods:
		return asposecells.SaveFormat_Ods
	case asposecells.FileFormatType_Ots:
		return asposecells.SaveFormat_Ots
	case asposecells.FileFormatType_Png:
		return asposecells.SaveFormat_Png
	case asposecells.FileFormatType_Pdf:
		return asposecells.SaveFormat_Pdf
	case asposecells.FileFormatType_Ppt,
		asposecells.FileFormatType_Pptx,
		asposecells.FileFormatType_Ppsm:
		return asposecells.SaveFormat_Pptx
	case asposecells.FileFormatType_Sxc:
		return asposecells.SaveFormat_Sxc
	case asposecells.FileFormatType_Svg:
		return asposecells.SaveFormat_Svg
	case asposecells.FileFormatType_SqlScript:
		return asposecells.SaveFormat_SqlScript
	case asposecells.FileFormatType_SpreadsheetML:
		return asposecells.SaveFormat_SpreadsheetML
	case asposecells.FileFormatType_Tsv:
		return asposecells.SaveFormat_Tsv
	case asposecells.FileFormatType_Tiff:
		return asposecells.SaveFormat_Tiff
	case asposecells.FileFormatType_Xlsm:
		return asposecells.SaveFormat_Xlsm
	case asposecells.FileFormatType_Xlsx:
		return asposecells.SaveFormat_Xlsx
	case asposecells.FileFormatType_Xlsb:
		return asposecells.SaveFormat_Xlsb
	case asposecells.FileFormatType_Xlam:
		return asposecells.SaveFormat_Xlam
	case asposecells.FileFormatType_Xlt:
		return asposecells.SaveFormat_Xlt
	case asposecells.FileFormatType_Xltx:
		return asposecells.SaveFormat_Xltx
	case asposecells.FileFormatType_Xml:
		return asposecells.SaveFormat_Xml
	case asposecells.FileFormatType_Xps:
		return asposecells.SaveFormat_Xps
	}

	return asposecells.SaveFormat_Auto
}
