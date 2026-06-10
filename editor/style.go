package editor

import asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"

// WithFontName creates a StyleAction that sets the font name (typeface) of the style.
//
// Parameters:
//   - fontName: The name of the font to apply (e.g., "Arial", "Times New Roman").
//
// Returns:
//   - StyleAction: A function that modifies the font name of the target style.
func WithFontName(fontName string) StyleAction {
	return func(style *asposecells.Style) error {
		font, _ := style.GetFont()
		font.SetName_String(fontName)
		return nil
	}
}

// WithFontSize creates a StyleAction that sets the font size of the style.
//
// Parameters:
//   - size: The font size in points.
//
// Returns:
//   - StyleAction: A function that modifies the font size of the target style.
func WithFontSize(size int) StyleAction {
	return func(style *asposecells.Style) error {
		font, _ := style.GetFont()
		font.SetSize(int32(size))
		return nil
	}
}

// WithFontIsBold creates a StyleAction that sets whether the font is bold.
//
// Parameters:
//   - value: true to make the font bold, false for normal weight.
//
// Returns:
//   - StyleAction: A function that modifies the bold property of the target style.
func WithFontIsBold(value bool) StyleAction {
	return func(style *asposecells.Style) error {
		font, _ := style.GetFont()
		return font.SetIsBold(value)
	}
}

// WithFontIsItalic creates a StyleAction that sets whether the font is italicized.
//
// Parameters:
//   - value: true to make the font italic, false for normal slant.
//
// Returns:
//   - StyleAction: A function that modifies the italic property of the target style.
func WithFontIsItalic(value bool) StyleAction {
	return func(style *asposecells.Style) error {
		font, _ := style.GetFont()
		return font.SetIsItalic(value)
	}
}

// WithFontIsSuperscript creates a StyleAction that sets whether the font is formatted as superscript.
//
// Parameters:
//   - value: true to format as superscript, false for normal baseline.
//
// Returns:
//   - StyleAction: A function that modifies the superscript property of the target style.
func WithFontIsSuperscript(value bool) StyleAction {
	return func(style *asposecells.Style) error {
		font, _ := style.GetFont()
		return font.SetIsSuperscript(value)
	}
}

// WithFontIsSubscript creates a StyleAction that sets whether the font is formatted as subscript.
//
// Parameters:
//   - value: true to format as subscript, false for normal baseline.
//
// Returns:
//   - StyleAction: A function that modifies the subscript property of the target style.
func WithFontIsSubscript(value bool) StyleAction {
	return func(style *asposecells.Style) error {
		font, _ := style.GetFont()
		return font.SetIsSubscript(value)
	}
}

// WithFontIsStrikeout creates a StyleAction that sets whether the font has a strikethrough line.
//
// Parameters:
//   - value: true to apply strikethrough, false to remove it.
//
// Returns:
//   - StyleAction: A function that modifies the strikeout property of the target style.
func WithFontIsStrikeout(value bool) StyleAction {
	return func(style *asposecells.Style) error {
		font, _ := style.GetFont()
		return font.SetIsStrikeout(value)
	}
}

// WithFontColor creates a StyleAction that sets the color of the font.
// This function uses a flexible parameter type to support multiple color representations.
//
// Parameters:
//   - color: An interface{} value representing the desired font color.
//     Supported types are resolved internally by the resolveColor helper
//     (e.g., hex strings like "#FF0000", or native Aspose.Color objects).
//
// Returns:
//   - StyleAction: A function that modifies the font color of the target style.
func WithFontColor(color interface{}) StyleAction {
	return func(style *asposecells.Style) error {
		v, err := resolveColor(color)
		if err == nil {
			font, err := style.GetFont()
			if err == nil {
				return font.SetColor(v)
			}
		}
		return err
	}
}

// WithFontUnderline creates a StyleAction that sets the underline style of the font.
// This function uses a flexible parameter type to support multiple underline representations.
//
// Parameters:
//   - value: An interface{} value representing the desired underline style.
//     Supported types are resolved internally by the resolveFontUnderline helper
//     (e.g., string names like "single", "double", or native FontUnderlineType enums).
//
// Returns:
//   - StyleAction: A function that modifies the font underline style of the target style.
func WithFontUnderline(value interface{}) StyleAction {
	return func(style *asposecells.Style) error {
		font, _ := style.GetFont()
		line, err := resolveFontUnderline(value)
		if err != nil {
			return err
		}
		return font.SetUnderline(line)
	}
}

// WithBackgroundColor creates a StyleAction that sets the background color of the style.
// Under the hood, this function automatically sets the fill pattern to Solid
// (BackgroundType_Solid) and applies the provided color as the foreground color.
//
// Parameters:
//   - color: An interface{} value representing the desired background color.
//     Supported types are resolved internally by the resolveColor helper
//     (e.g., hex strings like "#FFFF00", "red", or native Aspose.Color objects).
//
// Returns:
//   - StyleAction: A function that modifies the background fill of the target style.
func WithBackgroundColor(color interface{}) StyleAction {
	return func(style *asposecells.Style) error {
		v, err := resolveColor(color)
		if err == nil {
			style.SetForegroundColor(v)
			style.SetPattern(asposecells.BackgroundType_Solid)
			return nil
		}
		return err
	}
}

// WithHorizontalAlignment creates a StyleAction that sets the horizontal text alignment of the style.
//
// Parameters:
//   - alignment: An interface{} value representing the desired horizontal alignment.
//     Supported types are resolved internally by the resolveTextAlignmentType helper
//     (e.g., string names like "center", "left", or native TextAlignmentType enums).
//
// Returns:
//   - StyleAction: A function that modifies the horizontal alignment of the target style.
func WithHorizontalAlignment(alignment interface{}) StyleAction {
	return func(style *asposecells.Style) error {
		v, err := resolveTextAlignmentType(alignment)
		if err == nil {
			return style.SetHorizontalAlignment(v)
		}
		return err
	}
}

// WithVerticalAlignment creates a StyleAction that sets the vertical text alignment of the style.
//
// Parameters:
//   - alignment: An interface{} value representing the desired vertical alignment.
//     Supported types are resolved internally by the resolveTextAlignmentType helper
//     (e.g., string names like "top", "bottom", "center", or native TextAlignmentType enums).
//
// Returns:
//   - StyleAction: A function that modifies the vertical alignment of the target style.
func WithVerticalAlignment(alignment interface{}) StyleAction {
	return func(style *asposecells.Style) error {
		v, err := resolveTextAlignmentType(alignment)
		if err == nil {
			return style.SetVerticalAlignment(v)
		}
		return err
	}
}

// WithIsTextWrapped creates a StyleAction that enables or disables text wrapping within the cell.
//
// Parameters:
//   - value: true to enable text wrapping, false to disable it.
//
// Returns:
//   - StyleAction: A function that modifies the text wrapping property of the target style.
func WithIsTextWrapped(value bool) StyleAction {
	return func(style *asposecells.Style) error {
		style.SetIsTextWrapped(value)
		return nil
	}
}

// WithIsAlignmentApplied creates a StyleAction that explicitly marks the alignment properties
// as applied in this style. This is often required to ensure alignment overrides take effect
// when inheriting from default styles.
//
// Parameters:
//   - value: true to mark alignment properties as applied.
//
// Returns:
//   - StyleAction: A function that updates the style's applied flags.
func WithIsAlignmentApplied(value bool) StyleAction {
	return func(style *asposecells.Style) error {
		style.SetIsAlignmentApplied(value)
		return nil
	}
}

// WithIsBorderApplied creates a StyleAction that explicitly marks border properties
// as applied in this style.
//
// Parameters:
//   - value: true to mark border properties as applied.
//
// Returns:
//   - StyleAction: A function that updates the style's applied flags.
func WithIsBorderApplied(value bool) StyleAction {
	return func(style *asposecells.Style) error {
		style.SetIsBorderApplied(value)
		return nil
	}
}

// WithIsFillApplied creates a StyleAction that explicitly marks fill/background properties
// as applied in this style.
//
// Parameters:
//   - value: true to mark fill properties as applied.
//
// Returns:
//   - StyleAction: A function that updates the style's applied flags.
func WithIsFillApplied(value bool) StyleAction {
	return func(style *asposecells.Style) error {
		style.SetIsFillApplied(value)
		return nil
	}
}

// WithIsProtectionApplied creates a StyleAction that explicitly marks protection settings
// as applied in this style.
//
// Parameters:
//   - value: true to mark protection properties as applied.
//
// Returns:
//   - StyleAction: A function that updates the style's applied flags.
func WithIsProtectionApplied(value bool) StyleAction {
	return func(style *asposecells.Style) error {
		style.SetIsProtectionApplied(value)
		return nil
	}
}

// WithIndentLevel creates a StyleAction that sets the indentation level for the text within the cell.
//
// Parameters:
//   - value: The number of indent levels to apply (zero-based).
//
// Returns:
//   - StyleAction: A function that modifies the text indentation of the target style.
func WithIndentLevel(value int) StyleAction {
	return func(style *asposecells.Style) error {
		style.SetIndentLevel(int32(value))
		return nil
	}
}
