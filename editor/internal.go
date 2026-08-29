package editor

import (
	"encoding/hex"
	"fmt"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
	"strings"
)

func resolveWorksheet(wb *asposecells.Workbook, id interface{}) (*asposecells.Worksheet, error) {
	wss, err := wb.GetWorksheets()
	if err != nil {
		return nil, err
	}
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
		return nil, fmt.Errorf("invalid sheet identifier: %v", id)
	}
}
func resolveColor(value interface{}) (*asposecells.Color, error) {
	if strVal, ok := value.(string); ok {
		// Only treat 6/8-digit hex strings (RGB/RGBA) as colors,
		// otherwise a color name made of hex chars (e.g. "face") would be misread.
		h := strings.TrimPrefix(strVal, "#")
		if len(h) == 6 || len(h) == 8 {
			if _, err := hex.DecodeString(h); err == nil {
				return asposecells.Color_FromHex(h)
			}
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
