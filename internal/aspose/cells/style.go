package cells

import (
	engine "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/engine"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// CellStyle is a toolkit-native snapshot of a cell's most commonly read
// formatting: the font's name, size, weight, slant, and color, plus the number
// format. It deliberately holds plain Go values so callers never touch an
// engine Style, Font, or Color handle.
type CellStyle struct {
	// FontName is the font family name, e.g. "Calibri".
	FontName string
	// FontSize is the font size in points.
	FontSize int32
	// Bold and Italic report the font's weight and slant.
	Bold   bool
	Italic bool
	// FontColor is the font color as an "#RRGGBB" hex string. Color.ToHex
	// returns it without a leading "#", so one is added here.
	FontColor string
	// NumberFormat is the built-in number format id (0 = General). CustomFormat
	// is the format string when the cell uses a custom format; it is empty for a
	// built-in one.
	NumberFormat int32
	CustomFormat string
}

// ReadCellStyle reads cell's formatting into a CellStyle.
func ReadCellStyle(cell *asposecells.Cell) (CellStyle, error) {
	style, err := engine.Derive(cell.GetStyle())
	if err != nil {
		return CellStyle{}, err
	}
	font, err := engine.Derive(style.GetFont())
	if err != nil {
		return CellStyle{}, err
	}
	name, err := font.GetName()
	if err != nil {
		return CellStyle{}, err
	}
	size, err := font.GetSize()
	if err != nil {
		return CellStyle{}, err
	}
	bold, err := font.IsBold()
	if err != nil {
		return CellStyle{}, err
	}
	italic, err := font.IsItalic()
	if err != nil {
		return CellStyle{}, err
	}
	fontColor, err := font.GetColor()
	if err != nil {
		return CellStyle{}, err
	}
	hex, err := fontColor.ToHex(false)
	if err != nil {
		return CellStyle{}, err
	}
	number, err := style.GetNumber()
	if err != nil {
		return CellStyle{}, err
	}
	custom, err := style.GetCustom()
	if err != nil {
		return CellStyle{}, err
	}
	return CellStyle{
		FontName:     name,
		FontSize:     size,
		Bold:         bold,
		Italic:       italic,
		FontColor:    normalizeHexColor(hex),
		NumberFormat: number,
		CustomFormat: custom,
	}, nil
}

// normalizeHexColor gives Color.ToHex's bare "RRGGBB" a leading "#". The engine
// has been observed to include the "#" already, so both shapes are accepted.
func normalizeHexColor(hex string) string {
	if hex == "" || hex[0] == '#' {
		return hex
	}
	return "#" + hex
}
