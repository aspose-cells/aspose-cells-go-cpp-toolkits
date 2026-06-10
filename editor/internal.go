package editor

import (
	"encoding/hex"
	"fmt"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
	"strings"
)

func resolveWorksheet(wb *asposecells.Workbook, id interface{}) (*asposecells.Worksheet, error) {
	wss, _ := wb.GetWorksheets()
	switch v := id.(type) {
	case int:
		return wss.Get_Int(int32(v))
	case int32:
		return wss.Get_Int(int32(v))
	case int64:
		return wss.Get_Int(int32(v))
	case string:
		return wss.Get_String(v)
	default:
		println("default")
		return nil, fmt.Errorf("invalid sheet identifier: %v", id)
	}
	return nil, fmt.Errorf("invalid sheet identifier: %v", id)
}
func resolveColor(value interface{}) (*asposecells.Color, error) {
	if strVal, ok := value.(string); ok {
		_, err := hex.DecodeString(strVal)
		if err == nil {
			return asposecells.Color_FromHex(strVal)
		}
		return asposecells.Color_FromName(strVal)
	} else if intVal, ok := value.(int); ok {
		return asposecells.Color_FromArgb(int32(intVal))
	} else if colorVal, ok := value.(*asposecells.Color); ok {
		return colorVal, nil
	} else {
		return nil, fmt.Errorf("invalid color value: %v", value)
	}
}

func resolveFontUnderline(value interface{}) (asposecells.FontUnderlineType, error) {
	if enmuVal, ok := value.(asposecells.FontUnderlineType); ok {
		return enmuVal, nil
	}
	if strVal, ok := value.(string); ok {
		switch strings.ToLower(strVal) {
		case "none":
			return asposecells.FontUnderlineType_None, nil
		case "single":
			return asposecells.FontUnderlineType_Single, nil
		case "double":
			return asposecells.FontUnderlineType_Double, nil
		case "accounting":
			return asposecells.FontUnderlineType_Accounting, nil
		case "doubleaccounting":
			return asposecells.FontUnderlineType_DoubleAccounting, nil
		case "dash":
			return asposecells.FontUnderlineType_Dash, nil
		case "dashdotdotheavy":
			return asposecells.FontUnderlineType_DashDotDotHeavy, nil
		case "dashdotheavy":
			return asposecells.FontUnderlineType_DashDotHeavy, nil
		case "dashlong":
			return asposecells.FontUnderlineType_DashLong, nil
		case "dashlongheavy":
			return asposecells.FontUnderlineType_DashLongHeavy, nil
		case "dotdash":
			return asposecells.FontUnderlineType_DotDash, nil
		case "dotdotdash":
			return asposecells.FontUnderlineType_DotDotDash, nil
		case "dotted":
			return asposecells.FontUnderlineType_Dotted, nil
		case "dottedheavy":
			return asposecells.FontUnderlineType_DottedHeavy, nil
		case "heavy":
			return asposecells.FontUnderlineType_Heavy, nil
		case "wave":
			return asposecells.FontUnderlineType_Wave, nil
		case "wavydouble":
			return asposecells.FontUnderlineType_WavyDouble, nil
		case "wavyheavy":
			return asposecells.FontUnderlineType_WavyHeavy, nil
		case "words":
			return asposecells.FontUnderlineType_Words, nil
		default:
			return asposecells.FontUnderlineType_None, nil
		}
	}
	return asposecells.FontUnderlineType_None, nil
}
func resolveTextAlignmentType(value interface{}) (asposecells.TextAlignmentType, error) {
	if enumVal, ok := value.(asposecells.TextAlignmentType); ok {
		return enumVal, nil
	}

	if strVal, ok := value.(string); ok {
		switch strings.ToLower(strVal) {
		case "general":
			return asposecells.TextAlignmentType_General, nil
		case "bottom":
			return asposecells.TextAlignmentType_Bottom, nil
		case "center":
			return asposecells.TextAlignmentType_Center, nil
		case "centeracross":
			return asposecells.TextAlignmentType_CenterAcross, nil
		case "distributed":
			return asposecells.TextAlignmentType_Distributed, nil
		case "fill":
			return asposecells.TextAlignmentType_Fill, nil
		case "justify":
			return asposecells.TextAlignmentType_Justify, nil
		case "left":
			return asposecells.TextAlignmentType_Left, nil
		case "right":
			return asposecells.TextAlignmentType_Right, nil
		case "top":
			return asposecells.TextAlignmentType_Top, nil
		case "justifiedlow":
			return asposecells.TextAlignmentType_JustifiedLow, nil
		case "thaidistributed":
			return asposecells.TextAlignmentType_ThaiDistributed, nil
		default:
			return asposecells.TextAlignmentType_General, nil
		}
	}

	return asposecells.TextAlignmentType_General, nil
}

func fileFormatToSaveFormat(formatType asposecells.FileFormatType) asposecells.SaveFormat {
	if formatType == asposecells.FileFormatType_Azw3 {
		return asposecells.SaveFormat_Azw3
	}
	if formatType == asposecells.FileFormatType_Bmp {
		return asposecells.SaveFormat_Bmp
	}
	if formatType == asposecells.FileFormatType_Csv {
		return asposecells.SaveFormat_Csv
	}
	if formatType == asposecells.FileFormatType_Docx {
		return asposecells.SaveFormat_Docx
	}
	if formatType == asposecells.FileFormatType_Docm {
		return asposecells.SaveFormat_Docx
	}
	if formatType == asposecells.FileFormatType_Doc {
		return asposecells.SaveFormat_Docx
	}
	if formatType == asposecells.FileFormatType_Dotm {
		return asposecells.SaveFormat_Docx
	}
	if formatType == asposecells.FileFormatType_Dif {
		return asposecells.SaveFormat_Dif
	}
	if formatType == asposecells.FileFormatType_Dbf {
		return asposecells.SaveFormat_Dbf
	}
	if formatType == asposecells.FileFormatType_Epub {
		return asposecells.SaveFormat_Epub
	}
	if formatType == asposecells.FileFormatType_Emf {
		return asposecells.SaveFormat_Emf
	}
	if formatType == asposecells.FileFormatType_Excel97To2003 {
		return asposecells.SaveFormat_Excel97To2003
	}
	if formatType == asposecells.FileFormatType_Fods {
		return asposecells.SaveFormat_Fods
	}
	if formatType == asposecells.FileFormatType_Gif {
		return asposecells.SaveFormat_Gif
	}
	if formatType == asposecells.FileFormatType_Html {
		return asposecells.SaveFormat_Html
	}
	if formatType == asposecells.FileFormatType_MHtml {
		return asposecells.SaveFormat_MHtml
	}
	if formatType == asposecells.FileFormatType_Json {
		return asposecells.SaveFormat_Json
	}
	if formatType == asposecells.FileFormatType_Jpg {
		return asposecells.SaveFormat_Jpg
	}
	if formatType == asposecells.FileFormatType_Markdown {
		return asposecells.SaveFormat_Markdown
	}
	if formatType == asposecells.FileFormatType_Numbers35 {
		return asposecells.SaveFormat_Numbers
	}
	if formatType == asposecells.FileFormatType_Numbers35 {
		return asposecells.SaveFormat_Numbers
	}
	if formatType == asposecells.FileFormatType_Ods {
		return asposecells.SaveFormat_Ods
	}
	if formatType == asposecells.FileFormatType_Ots {
		return asposecells.SaveFormat_Ots
	}
	if formatType == asposecells.FileFormatType_Png {
		return asposecells.SaveFormat_Png
	}
	if formatType == asposecells.FileFormatType_Pdf {
		return asposecells.SaveFormat_Pdf
	}
	if formatType == asposecells.FileFormatType_Ppt {
		return asposecells.SaveFormat_Pptx
	}
	if formatType == asposecells.FileFormatType_Pptx {
		return asposecells.SaveFormat_Pptx
	}
	if formatType == asposecells.FileFormatType_Ppsm {
		return asposecells.SaveFormat_Pptx
	}
	if formatType == asposecells.FileFormatType_Rtf {
		return asposecells.SaveFormat_Docx
	}
	if formatType == asposecells.FileFormatType_Rtf {
		return asposecells.SaveFormat_Docx
	}
	if formatType == asposecells.FileFormatType_Sxc {
		return asposecells.SaveFormat_Sxc
	}
	if formatType == asposecells.FileFormatType_Svg {
		return asposecells.SaveFormat_Svg
	}
	if formatType == asposecells.FileFormatType_SqlScript {
		return asposecells.SaveFormat_SqlScript
	}
	if formatType == asposecells.FileFormatType_SpreadsheetML {
		return asposecells.SaveFormat_SpreadsheetML
	}
	if formatType == asposecells.FileFormatType_Tsv {
		return asposecells.SaveFormat_Tsv
	}
	if formatType == asposecells.FileFormatType_Tiff {
		return asposecells.SaveFormat_Tiff
	}
	if formatType == asposecells.FileFormatType_Tsv {
		return asposecells.SaveFormat_Tsv
	}
	if formatType == asposecells.FileFormatType_Xlsm {
		return asposecells.SaveFormat_Xlsm
	}
	if formatType == asposecells.FileFormatType_Xlsx {
		return asposecells.SaveFormat_Xlsx
	}
	if formatType == asposecells.FileFormatType_Xlsb {
		return asposecells.SaveFormat_Xlsb
	}
	if formatType == asposecells.FileFormatType_Xlam {
		return asposecells.SaveFormat_Xlam
	}
	if formatType == asposecells.FileFormatType_Xlt {
		return asposecells.SaveFormat_Xlt
	}
	if formatType == asposecells.FileFormatType_Xltx {
		return asposecells.SaveFormat_Xltx
	}
	if formatType == asposecells.FileFormatType_Xml {
		return asposecells.SaveFormat_Xml
	}
	if formatType == asposecells.FileFormatType_Xps {
		return asposecells.SaveFormat_Xps
	}

	return asposecells.SaveFormat_Auto
}
