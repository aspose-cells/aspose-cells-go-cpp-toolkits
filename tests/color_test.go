package tests

import (
	"bytes"
	"errors"
	"image/color"
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	asposecolor "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/color"
	engine "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/engine"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/docx"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/pcl"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/pdf"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/pptx"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/xps"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// argb packs a channel quadruple the way the toolkit and the engine write it.
func argb(a, r, g, b uint8) uint32 {
	return uint32(a)<<24 | uint32(r)<<16 | uint32(g)<<8 | uint32(b)
}

// colorArgb reads an engine Color back as ARGB.
func colorArgb(t *testing.T, c *asposecells.Color) uint32 {
	t.Helper()
	a, err := c.Get_Color_A()
	if err != nil {
		t.Fatalf("Get_Color_A error: %v", err)
	}
	r, err := c.Get_Color_R()
	if err != nil {
		t.Fatalf("Get_Color_R error: %v", err)
	}
	g, err := c.Get_Color_G()
	if err != nil {
		t.Fatalf("Get_Color_G error: %v", err)
	}
	b, err := c.Get_Color_B()
	if err != nil {
		t.Fatalf("Get_Color_B error: %v", err)
	}
	return argb(a, r, g, b)
}

// resolveArgb resolves a color and reads the result back out of the engine.
func resolveArgb(t *testing.T, value interface{}) uint32 {
	t.Helper()
	c, owned, err := asposecolor.Resolve(value)
	if err != nil {
		t.Fatalf("Resolve(%v) error: %v", value, err)
	}
	if !owned {
		t.Fatalf("Resolve(%v) reported the color as the caller's, want the toolkit's", value)
	}
	defer asposecells.DeleteColor(c)
	return colorArgb(t, c)
}

// TestColorResolveGoColor pins the conversion from a Go image/color value to the
// engine's straight-alpha ARGB. The translucent cases are the point of it: every
// color.Color reports alpha-premultiplied channels, so a translucent value has to
// have its alpha divided back out on the way in. Passing the premultiplied
// channels straight through would turn "full red at half alpha" into "half red at
// half alpha", and the engine would faithfully paint the darker color.
func TestColorResolveGoColor(t *testing.T) {
	cases := []struct {
		name  string
		value color.Color
		want  uint32
	}{
		{"RGBA opaque", color.RGBA{R: 0x11, G: 0x22, B: 0x33, A: 0xFF}, argb(0xFF, 0x11, 0x22, 0x33)},
		{"NRGBA opaque", color.NRGBA{R: 0x11, G: 0x22, B: 0x33, A: 0xFF}, argb(0xFF, 0x11, 0x22, 0x33)},
		{"RGBA and NRGBA agree", color.RGBA{R: 0x11, G: 0x22, B: 0x33, A: 0xFF}, argb(0xFF, 0x11, 0x22, 0x33)},
		{"NRGBA translucent", color.NRGBA{R: 0xFF, A: 0x80}, argb(0x80, 0xFF, 0x00, 0x00)},
		{"RGBA premultiplied translucent", color.RGBA{R: 0x80, A: 0x80}, argb(0x80, 0xFF, 0x00, 0x00)},
		// color.RGBA is documented as already premultiplied, so R=0xFF with A=0x80
		// is out of gamut. It is also how people spell "half-transparent red", and
		// dividing the alpha out clamps it to the full red they meant rather than
		// wrapping around to an unrelated channel value.
		{"RGBA out of gamut clamps", color.RGBA{R: 0xFF, A: 0x80}, argb(0x80, 0xFF, 0x00, 0x00)},
		{"NRGBA fully transparent", color.NRGBA{R: 0xFF, A: 0x00}, argb(0x00, 0x00, 0x00, 0x00)},
		{"Gray", color.Gray{Y: 0x80}, argb(0xFF, 0x80, 0x80, 0x80)},
		{"named black", color.Black, argb(0xFF, 0x00, 0x00, 0x00)},
		{"named white", color.White, argb(0xFF, 0xFF, 0xFF, 0xFF)},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			if got := resolveArgb(t, tc.value); got != tc.want {
				t.Errorf("Resolve(%#v) => %#08x, want %#08x", tc.value, got, tc.want)
			}
		})
	}
}

// TestColorResolveForms checks that the four spellings of a color name the same
// thing, which is what lets every color-taking option accept the same set.
func TestColorResolveForms(t *testing.T) {
	cases := []struct {
		name  string
		value interface{}
		want  uint32
	}{
		{"hex string", "#336699", argb(0xFF, 0x33, 0x66, 0x99)},
		{"hex string without hash", "336699", argb(0xFF, 0x33, 0x66, 0x99)},
		{"hex string with alpha", "33669980", argb(0x80, 0x33, 0x66, 0x99)},
		{"argb int", 0xFF336699, argb(0xFF, 0x33, 0x66, 0x99)},
		{"color name", "red", argb(0xFF, 0xFF, 0x00, 0x00)},
		{"color name with spaces", "Light Sea Green", argb(0xFF, 0x20, 0xB2, 0xAA)},
		{"color name case-insensitive", "LIGHTSEAGREEN", argb(0xFF, 0x20, 0xB2, 0xAA)},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			if got := resolveArgb(t, tc.value); got != tc.want {
				t.Errorf("Resolve(%v) => %#08x, want %#08x", tc.value, got, tc.want)
			}
		})
	}
}

// TestColorResolveRejects checks that a value the toolkit cannot turn into a
// color is reported rather than quietly leaving the default in place. The
// unknown name in particular must not reach the engine's own Color_FromName,
// whose C++ throw crosses cgo and kills the process.
func TestColorResolveRejects(t *testing.T) {
	for _, value := range []interface{}{"notacolor", "#GGGGGG", 3.14, struct{}{}} {
		t.Run("", func(t *testing.T) {
			c, owned, err := asposecolor.Resolve(value)
			if !errors.Is(err, toolkiterrors.ErrInvalidColor) {
				t.Fatalf("Resolve(%v) error = %v, want ErrInvalidColor", value, err)
			}
			if c != nil || owned {
				t.Errorf("Resolve(%v) returned a color alongside the error", value)
			}
		})
	}
}

// TestColorResolveOwnership checks that Resolve reports who has to release the
// engine Color. The engine's Color carries no finalizer, so a color the toolkit
// builds leaks permanently unless somebody deletes it - and a color the caller
// built must not be deleted by anyone else, or the caller holds a freed object.
func TestColorResolveOwnership(t *testing.T) {
	callerArgb := uint32(0xFF336699) // out of range as an int32 constant
	caller, err := asposecells.Color_FromArgb(int32(callerArgb))
	if err != nil {
		t.Fatalf("Color_FromArgb error: %v", err)
	}
	defer asposecells.DeleteColor(caller)

	got, owned, err := asposecolor.Resolve(caller)
	if err != nil {
		t.Fatalf("Resolve(engine color) error: %v", err)
	}
	if got != caller {
		t.Error("Resolve(engine color) returned a different color")
	}
	if owned {
		t.Error("Resolve(engine color) claimed a color the caller built as the toolkit's")
	}

	// Every other form is built by the toolkit, so the caller of Resolve has to
	// release it.
	for _, value := range []interface{}{"#336699", "red", 0xFF336699, color.RGBA{A: 0xFF}} {
		c, owned, err := asposecolor.Resolve(value)
		if err != nil {
			t.Fatalf("Resolve(%v) error: %v", value, err)
		}
		if !owned {
			t.Errorf("Resolve(%v) reported the toolkit's own color as the caller's", value)
		}
		asposecells.DeleteColor(c)
	}
}

// TestEditorFontColorFromGoColor covers the editor's end of the shared color
// vocabulary. Font colors are stored without an alpha channel - Excel font
// colors are opaque - so the assertions are on the channels the engine keeps,
// and the translucent cases assert that the color arrives at full strength
// rather than darkened by its own alpha.
func TestEditorFontColorFromGoColor(t *testing.T) {
	cases := []struct {
		name  string
		value color.Color
		want  uint32
	}{
		{"opaque RGB", color.NRGBA{R: 0x11, G: 0x22, B: 0x33, A: 0xFF}, argb(0xFF, 0x11, 0x22, 0x33)},
		{"translucent NRGBA keeps full strength", color.NRGBA{R: 0xFF, A: 0x80}, argb(0xFF, 0xFF, 0x00, 0x00)},
		{"translucent RGBA keeps full strength", color.RGBA{R: 0xFF, A: 0x80}, argb(0xFF, 0xFF, 0x00, 0x00)},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			style := newTestStyle(t)
			if err := editor.WithFontColor(tc.value)(style); err != nil {
				t.Fatalf("WithFontColor(%#v) error: %v", tc.value, err)
			}
			font, err := engine.Derive(style.GetFont())
			if err != nil {
				t.Fatalf("Derive font error: %v", err)
			}
			got, err := font.GetColor()
			if err != nil {
				t.Fatalf("GetColor error: %v", err)
			}
			if got == nil {
				t.Fatal("GetColor returned nil")
			}
			if argbValue := colorArgb(t, got); argbValue != tc.want {
				t.Errorf("WithFontColor(%#v) => %#08x, want %#08x", tc.value, argbValue, tc.want)
			}
		})
	}
}

// gridlineSavers is the five save options that take a gridline color, each
// wrapped so the tables below can drive them alike.
func gridlineSavers(t *testing.T) map[string]func(interface{}) ([]byte, error) {
	t.Helper()
	return map[string]func(interface{}) ([]byte, error){
		"docx": func(v interface{}) ([]byte, error) { return docx.New(docx.WithGridlineColor(v)).Apply(p2Workbook(t)) },
		"pcl":  func(v interface{}) ([]byte, error) { return pcl.New(pcl.WithGridlineColor(v)).Apply(p2Workbook(t)) },
		"pdf":  func(v interface{}) ([]byte, error) { return pdf.New(pdf.WithGridlineColor(v)).Apply(p2Workbook(t)) },
		"pptx": func(v interface{}) ([]byte, error) { return pptx.New(pptx.WithGridlineColor(v)).Apply(p2Workbook(t)) },
		"xps":  func(v interface{}) ([]byte, error) { return xps.New(xps.WithGridlineColor(v)).Apply(p2Workbook(t)) },
	}
}

// TestGridlineColorGoColor drives every gridline-color option with a Go
// image/color value, the form the toolkit added on top of the engine's *Color.
// The assertion is weak on purpose - a page's gridline color does not read back
// out of the produced document - but it covers the conversion, the engine call
// and the release of the color the toolkit built: a double free or a
// use-after-free here takes the whole test binary down, not just an assertion.
func TestGridlineColorGoColor(t *testing.T) {
	for name, save := range gridlineSavers(t) {
		t.Run(name, func(t *testing.T) {
			out, err := save(color.NRGBA{R: 0x33, G: 0x66, B: 0x99, A: 0xFF})
			if err != nil {
				t.Fatalf("save with a Go color: %v", err)
			}
			if len(out) == 0 {
				t.Fatal("save with a Go color produced no output")
			}
			// Saving twice builds and releases a second color, so a color freed
			// twice would surface here.
			if _, err := save(color.NRGBA{R: 0x33, G: 0x66, B: 0x99, A: 0xFF}); err != nil {
				t.Fatalf("second save with a Go color: %v", err)
			}
		})
	}
}

// TestGridlineColorForms checks the remaining spellings on one of the five, and
// that an unusable value is reported before any output is produced.
func TestGridlineColorForms(t *testing.T) {
	forms := []struct {
		name  string
		value interface{}
	}{
		{"hex string", "#336699"},
		{"hex string with alpha", "#33669980"},
		{"color name", "Light Sea Green"},
		{"argb int", 0xFF336699},
	}
	for _, form := range forms {
		t.Run(form.name, func(t *testing.T) {
			out, err := pdf.New(pdf.WithGridlineColor(form.value)).Apply(p2Workbook(t))
			if err != nil {
				t.Fatalf("pdf with gridline color %v: %v", form.value, err)
			}
			if !bytes.HasPrefix(out, []byte("%PDF-")) {
				t.Errorf("pdf output does not start with %%PDF- (got %d bytes)", len(out))
			}
		})
	}

	for _, value := range []interface{}{"notacolor", 3.14} {
		t.Run("rejected", func(t *testing.T) {
			out, err := pdf.New(pdf.WithGridlineColor(value)).Apply(p2Workbook(t))
			if !errors.Is(err, toolkiterrors.ErrInvalidColor) {
				t.Fatalf("pdf with gridline color %v: error = %v, want ErrInvalidColor", value, err)
			}
			if out != nil {
				t.Errorf("pdf with gridline color %v produced %d bytes alongside the error", value, len(out))
			}
		})
	}
}

// TestGridlineColorCallerColorNotReleased passes an engine *Color the caller
// owns to a save option and then uses it again. The toolkit has to recognize
// that color as the caller's and leave it alone - the explicit release that
// keeps the toolkit's own colors from leaking is the same one that would leave
// the caller holding a freed object.
func TestGridlineColorCallerColorNotReleased(t *testing.T) {
	callerArgb := uint32(0xFF336699) // out of range as an int32 constant
	caller, err := asposecells.Color_FromArgb(int32(callerArgb))
	if err != nil {
		t.Fatalf("Color_FromArgb error: %v", err)
	}
	defer asposecells.DeleteColor(caller)

	for i := 0; i < 2; i++ {
		if _, err := pdf.New(pdf.WithGridlineColor(caller)).Apply(p2Workbook(t)); err != nil {
			t.Fatalf("save %d with a caller-owned color: %v", i+1, err)
		}
	}
	if got := colorArgb(t, caller); got != callerArgb {
		t.Errorf("caller-owned color reads %#08x after saving, want %#08x", got, callerArgb)
	}
}
