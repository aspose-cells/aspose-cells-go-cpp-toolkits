// Package color resolves a caller-supplied color into the engine's Color, so
// that neither editor nor the save options have to name that type in a public
// signature.
//
// One color vocabulary serves the whole toolkit. Resolve accepts a Go
// image/color value, a hex string, one of the engine's own color names, or an
// ARGB integer, and is the single place that decides how a color is spelled —
// so editor.WithFontColor and pdf.WithGridlineColor accept the same things.
package color

import (
	"encoding/hex"
	"fmt"
	stdcolor "image/color"
	"strings"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	enums "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/enums"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// Resolve turns a caller-supplied color into an engine Color.
//
// Accepted forms are a Go color.Color (color.RGBA, color.NRGBA, color.Gray,
// color.Black, …), a hex string ("#RRGGBB" / "#RRGGBBAA", with or without the
// "#"), one of the engine's color names ("red", "Light Sea Green", matched case-
// and punctuation-insensitively), and an ARGB int. Anything else is
// ErrInvalidColor.
//
// The second return reports whether the toolkit created the Color. The engine's
// Color carries no finalizer — NewColor, Color_FromArgb and Color_FromHex only
// allocate — so the only way a toolkit-created Color is ever released is an
// explicit DeleteColor, and the caller of Resolve is the one that has to do it.
// A Color passed in as a *asposecells.Color is the caller's and must not be
// freed here or by anyone else.
func Resolve(value interface{}) (*asposecells.Color, bool, error) {
	switch v := value.(type) {
	case string:
		// Only treat 6/8-digit hex strings (RGB/RGBA) as colors, otherwise a
		// color name made of hex chars (e.g. "face") would be misread.
		h := strings.TrimPrefix(v, "#")
		if len(h) == 6 || len(h) == 8 {
			if _, err := hex.DecodeString(h); err == nil {
				c, err := asposecells.Color_FromHex(engineHexOrder(h))
				if err != nil {
					return nil, false, fmt.Errorf("color %q: %w", v, err)
				}
				return c, true, nil
			}
		}
		ctor, ok := namedColorConstructors[enums.Normalize(v)]
		if !ok {
			return nil, false, fmt.Errorf("color name %q is not a recognized color: %w", v, toolkiterrors.ErrInvalidColor)
		}
		c, err := ctor()
		if err != nil {
			return nil, false, fmt.Errorf("color %q: %w", v, err)
		}
		return c, true, nil
	case int:
		c, err := asposecells.Color_FromArgb(int32(v))
		if err != nil {
			return nil, false, fmt.Errorf("color %d: %w", v, err)
		}
		return c, true, nil
	case stdcolor.Color:
		c, err := fromGoColor(v)
		if err != nil {
			return nil, false, err
		}
		return c, true, nil
	case *asposecells.Color:
		return v, false, nil
	default:
		return nil, false, fmt.Errorf("invalid color value %v: %w", value, toolkiterrors.ErrInvalidColor)
	}
}

// engineHexOrder puts an eight-digit hex color into the order the engine reads.
//
// The engine reads "#AARRGGBB" - alpha first, the Java and Android spelling -
// while the toolkit's color vocabulary documents "#RRGGBBAA", the order CSS
// uses and the one the alpha-at-the-end spelling implies. The alpha is moved to
// the front so that a caller who writes the documented form gets the color they
// asked for: reading "33669980" as the engine does turns an 80/255 opaque teal
// into a 33/255 opaque one. Six-digit strings carry no alpha and need no
// reordering.
func engineHexOrder(h string) string {
	if len(h) == 8 {
		return h[6:8] + h[0:6]
	}
	return h
}

// fromGoColor converts a Go color to the engine's ARGB form.
//
// Every color.Color reports alpha-premultiplied 16-bit channels, while a
// document color is straight (non-premultiplied) alpha, so the alpha is divided
// back out on the way in. Without that step a translucent color would come out
// darkened by its own alpha: color.NRGBA{R: 0xFF, A: 0x80} would reach the
// engine as 50% red at 50% alpha, painting a half-strength color over a
// half-transparent one.
//
// The engine's Color_FromArgb reports an argument it refuses, so a value it will
// not take surfaces as an error instead of becoming some other color.
func fromGoColor(value stdcolor.Color) (*asposecells.Color, error) {
	// RGBA returns 16-bit channels; the engine takes 8 bits per channel.
	r16, g16, b16, a16 := value.RGBA()
	argb := uint32(a16>>8) << 24
	if a16 != 0 {
		// A fully transparent color has no channel to recover, and dividing by
		// its alpha would be a division by zero.
		argb |= unpremultiply(r16, a16)>>8<<16 | unpremultiply(g16, a16)>>8<<8 | unpremultiply(b16, a16)>>8
	}
	c, err := asposecells.Color_FromArgb(int32(argb))
	if err != nil {
		return nil, fmt.Errorf("color %v: %w", value, err)
	}
	return c, nil
}

// unpremultiply divides the alpha back out of a premultiplied 16-bit channel.
//
// A premultiplied channel cannot exceed its alpha, but a caller can still write
// one that does: color.RGBA{R: 0xFF, A: 0x80} spells "full red at half alpha"
// the way people reach for it, even though color.RGBA is documented as already
// premultiplied. Dividing that out yields more than a channel can hold, so the
// result is clamped to full strength rather than wrapping around to an
// unrelated color.
func unpremultiply(channel, alpha uint32) uint32 {
	v := channel * 0xFFFF / alpha
	if v > 0xFFFF {
		return 0xFFFF
	}
	return v
}
