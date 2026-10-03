package editor

import (
	"encoding/hex"
	"fmt"
	"strings"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	engine "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/engine"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

func resolveWorksheet(wb *asposecells.Workbook, id interface{}) (*asposecells.Worksheet, error) {
	wss, err := engine.Derive(wb.GetWorksheets())
	if err != nil {
		return nil, err
	}
	switch v := id.(type) {
	case int:
		return cells.SheetByIndex(wb, v)
	case int32:
		return cells.SheetByIndex(wb, int(v))
	case int64:
		return cells.SheetByIndex(wb, int(v))
	case string:
		return cells.WorksheetByName(wss, v)
	default:
		return nil, fmt.Errorf("invalid sheet identifier %v: %w", id, toolkiterrors.ErrInvalidSheetID)
	}
}

// resolveColor turns a caller-supplied color into an engine Color. Accepted
// forms are a hex string ("#RRGGBB" / "#RRGGBBAA", with or without the "#"), a
// color name ("red", "Light Sea Green"), an ARGB int, and an engine *Color.
//
// A name is looked up in namedColorConstructors and never handed to the
// engine's Color_FromName, which throws an uncaught C++ exception — and so
// terminates the process — for a name it does not recognize. Any name outside
// the table is reported as ErrInvalidColor.
func resolveColor(value interface{}) (*asposecells.Color, error) {
	switch v := value.(type) {
	case string:
		// Only treat 6/8-digit hex strings (RGB/RGBA) as colors, otherwise a
		// color name made of hex chars (e.g. "face") would be misread.
		h := strings.TrimPrefix(v, "#")
		if len(h) == 6 || len(h) == 8 {
			if _, err := hex.DecodeString(h); err == nil {
				c, err := asposecells.Color_FromHex(h)
				if err != nil {
					return nil, fmt.Errorf("color %q: %w", v, err)
				}
				return c, nil
			}
		}
		key := normalizeEnumName(v)
		ctor, ok := namedColorConstructors[key]
		if !ok {
			return nil, fmt.Errorf("color name %q is not a recognized color: %w", v, toolkiterrors.ErrInvalidColor)
		}
		c, err := ctor()
		if err != nil {
			return nil, fmt.Errorf("color %q: %w", v, err)
		}
		return c, nil
	case int:
		c, err := asposecells.Color_FromArgb(int32(v))
		if err != nil {
			return nil, fmt.Errorf("color %d: %w", v, err)
		}
		return c, nil
	case *asposecells.Color:
		return v, nil
	default:
		return nil, fmt.Errorf("invalid color value %v: %w", value, toolkiterrors.ErrInvalidColor)
	}
}

// normalizeEnumName reduces an enum name to its lookup key: lower case with
// every non-alphanumeric removed, so "Light Sea Green", "lightseagreen" and
// "Light-Sea-Green" all resolve alike, as do "double accounting" and
// "DoubleAccounting".
func normalizeEnumName(name string) string {
	var b strings.Builder
	b.Grow(len(name))
	for _, r := range name {
		switch {
		case r >= 'a' && r <= 'z', r >= '0' && r <= '9':
			b.WriteRune(r)
		case r >= 'A' && r <= 'Z':
			b.WriteRune(r + ('a' - 'A'))
		}
	}
	return b.String()
}

// fontUnderlineByName maps a normalized underline name to the engine enum.
var fontUnderlineByName = map[string]asposecells.FontUnderlineType{
	"none":             asposecells.FontUnderlineType_None,
	"single":           asposecells.FontUnderlineType_Single,
	"double":           asposecells.FontUnderlineType_Double,
	"accounting":       asposecells.FontUnderlineType_Accounting,
	"doubleaccounting": asposecells.FontUnderlineType_DoubleAccounting,
	"dash":             asposecells.FontUnderlineType_Dash,
	"dashdotdotheavy":  asposecells.FontUnderlineType_DashDotDotHeavy,
	"dashdotheavy":     asposecells.FontUnderlineType_DashDotHeavy,
	"dashlong":         asposecells.FontUnderlineType_DashLong,
	"dashlongheavy":    asposecells.FontUnderlineType_DashLongHeavy,
	"dotdash":          asposecells.FontUnderlineType_DotDash,
	"dotdotdash":       asposecells.FontUnderlineType_DotDotDash,
	"dotted":           asposecells.FontUnderlineType_Dotted,
	"dottedheavy":      asposecells.FontUnderlineType_DottedHeavy,
	"heavy":            asposecells.FontUnderlineType_Heavy,
	"wave":             asposecells.FontUnderlineType_Wave,
	"wavydouble":       asposecells.FontUnderlineType_WavyDouble,
	"wavyheavy":        asposecells.FontUnderlineType_WavyHeavy,
	"words":            asposecells.FontUnderlineType_Words,
}

// resolveFontUnderline resolves an underline style from a name (matched case-
// and punctuation-insensitively, so "DoubleAccounting" and "double accounting"
// both work) or an engine FontUnderlineType. An unrecognized name is an error
// rather than a silent fall back to None, which would silently drop an
// underline the caller asked for.
func resolveFontUnderline(value interface{}) (asposecells.FontUnderlineType, error) {
	switch v := value.(type) {
	case asposecells.FontUnderlineType:
		return v, nil
	case string:
		if line, ok := fontUnderlineByName[normalizeEnumName(v)]; ok {
			return line, nil
		}
		return asposecells.FontUnderlineType_None,
			fmt.Errorf("font underline %q is not a recognized underline style: %w", v, toolkiterrors.ErrInvalidFontUnderline)
	default:
		return asposecells.FontUnderlineType_None,
			fmt.Errorf("invalid font underline value %v: %w", value, toolkiterrors.ErrInvalidFontUnderline)
	}
}

// textAlignmentTypeByName maps a normalized alignment name to the engine enum.
// Horizontal and vertical alignment share one engine enum, so one table serves
// both; the names that only make sense for one axis are harmless on the other
// (the engine accepts them, they simply have no visible effect).
var textAlignmentTypeByName = map[string]asposecells.TextAlignmentType{
	"general":         asposecells.TextAlignmentType_General,
	"bottom":          asposecells.TextAlignmentType_Bottom,
	"center":          asposecells.TextAlignmentType_Center,
	"centeracross":    asposecells.TextAlignmentType_CenterAcross,
	"distributed":     asposecells.TextAlignmentType_Distributed,
	"fill":            asposecells.TextAlignmentType_Fill,
	"justify":         asposecells.TextAlignmentType_Justify,
	"left":            asposecells.TextAlignmentType_Left,
	"right":           asposecells.TextAlignmentType_Right,
	"top":             asposecells.TextAlignmentType_Top,
	"justifiedlow":    asposecells.TextAlignmentType_JustifiedLow,
	"thaidistributed": asposecells.TextAlignmentType_ThaiDistributed,
}

// resolveTextAlignmentType resolves an alignment from a name (matched case- and
// punctuation-insensitively) or an engine TextAlignmentType. An unrecognized
// name is an error rather than a silent fall back to General, which would
// silently discard an alignment the caller asked for.
func resolveTextAlignmentType(value interface{}) (asposecells.TextAlignmentType, error) {
	switch v := value.(type) {
	case asposecells.TextAlignmentType:
		return v, nil
	case string:
		if a, ok := textAlignmentTypeByName[normalizeEnumName(v)]; ok {
			return a, nil
		}
		return asposecells.TextAlignmentType_General,
			fmt.Errorf("text alignment %q is not a recognized alignment: %w", v, toolkiterrors.ErrInvalidTextAlignment)
	default:
		return asposecells.TextAlignmentType_General,
			fmt.Errorf("invalid text alignment value %v: %w", value, toolkiterrors.ErrInvalidTextAlignment)
	}
}
