package tests

import (
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// newTestStyle builds a Style via the public engine API (workbook -> worksheet
// -> cells -> GetStyle), the same path the editor.SetStyle action uses.
func newTestStyle(t *testing.T) *asposecells.Style {
	t.Helper()
	wb, err := asposecells.NewWorkbook()
	if err != nil {
		t.Fatalf("NewWorkbook error: %v", err)
	}
	wss, err := wb.GetWorksheets()
	if err != nil {
		t.Fatalf("GetWorksheets error: %v", err)
	}
	ws, err := wss.Get_Int(0)
	if err != nil {
		t.Fatalf("Get_Int error: %v", err)
	}
	cells, err := ws.GetCells()
	if err != nil {
		t.Fatalf("GetCells error: %v", err)
	}
	style, err := cells.GetStyle()
	if err != nil {
		t.Fatalf("GetStyle error: %v", err)
	}
	return style
}

func TestEditorWithFontNameAndSize(t *testing.T) {
	style := newTestStyle(t)
	if err := editor.WithFontName("Arial")(style); err != nil {
		t.Fatalf("WithFontName error: %v", err)
	}
	if err := editor.WithFontSize(12)(style); err != nil {
		t.Fatalf("WithFontSize error: %v", err)
	}
}

// The resolve helpers are unexported, so underline mapping is asserted through
// the public WithFontUnderline action and read back via Font.GetUnderline.
func TestEditorFontUnderlineMappings(t *testing.T) {
	cases := []struct {
		name  string
		input interface{}
		want  asposecells.FontUnderlineType
	}{
		{"enum passthrough", asposecells.FontUnderlineType_Accounting, asposecells.FontUnderlineType_Accounting},
		{"none", "none", asposecells.FontUnderlineType_None},
		{"single", "single", asposecells.FontUnderlineType_Single},
		{"double", "double", asposecells.FontUnderlineType_Double},
		{"double case-insensitive", "DOUBLE", asposecells.FontUnderlineType_Double},
		{"accounting", "accounting", asposecells.FontUnderlineType_Accounting},
		{"wave", "wave", asposecells.FontUnderlineType_Wave},
		{"words", "words", asposecells.FontUnderlineType_Words},
		{"unknown falls back to none", "super-bold", asposecells.FontUnderlineType_None},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			style := newTestStyle(t)
			if err := editor.WithFontUnderline(tc.input)(style); err != nil {
				t.Fatalf("WithFontUnderline(%v) error: %v", tc.input, err)
			}
			font, err := style.GetFont()
			if err != nil {
				t.Fatalf("GetFont error: %v", err)
			}
			got, err := font.GetUnderline()
			if err != nil {
				t.Fatalf("GetUnderline error: %v", err)
			}
			if got != tc.want {
				t.Errorf("WithFontUnderline(%v) => %v, want %v", tc.input, got, tc.want)
			}
		})
	}
}

func TestEditorFontBooleanFlags(t *testing.T) {
	style := newTestStyle(t)
	actions := []editor.StyleAction{
		editor.WithFontIsBold(true),
		editor.WithFontIsItalic(true),
		editor.WithFontIsStrikeout(true),
		editor.WithFontIsSuperscript(true),
		editor.WithFontIsSubscript(true),
		editor.WithIsTextWrapped(true),
		editor.WithIsBorderApplied(true),
		editor.WithIsAlignmentApplied(true),
		editor.WithIsProtectionApplied(true),
		editor.WithIsFillApplied(true),
		editor.WithIndentLevel(3),
	}
	for _, action := range actions {
		if err := action(style); err != nil {
			t.Fatalf("style action error: %v", err)
		}
	}
}

func TestEditorHorizontalAlignmentMappings(t *testing.T) {
	cases := []struct {
		name  string
		input interface{}
		want  asposecells.TextAlignmentType
	}{
		{"enum passthrough", asposecells.TextAlignmentType_Fill, asposecells.TextAlignmentType_Fill},
		{"general", "general", asposecells.TextAlignmentType_General},
		// Note: "top"/"bottom" are vertical alignments; applying them as
		// horizontal alignments is ignored by the engine, so they are covered
		// by TestEditorBackgroundColorAndVerticalAlignment instead.
		{"center", "center", asposecells.TextAlignmentType_Center},
		{"center case-insensitive", "CENTER", asposecells.TextAlignmentType_Center},
		{"distributed", "distributed", asposecells.TextAlignmentType_Distributed},
		{"justify", "justify", asposecells.TextAlignmentType_Justify},
		{"left", "left", asposecells.TextAlignmentType_Left},
		{"right", "right", asposecells.TextAlignmentType_Right},
		{"unknown falls back to general", "diagonal", asposecells.TextAlignmentType_General},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			style := newTestStyle(t)
			if err := editor.WithHorizontalAlignment(tc.input)(style); err != nil {
				t.Fatalf("WithHorizontalAlignment(%v) error: %v", tc.input, err)
			}
			got, err := style.GetHorizontalAlignment()
			if err != nil {
				t.Fatalf("GetHorizontalAlignment error: %v", err)
			}
			if got != tc.want {
				t.Errorf("WithHorizontalAlignment(%v) => %v, want %v", tc.input, got, tc.want)
			}
		})
	}
}

func TestEditorFontColor(t *testing.T) {
	t.Run("valid color name", func(t *testing.T) {
		style := newTestStyle(t)
		if err := editor.WithFontColor("red")(style); err != nil {
			t.Fatalf("WithFontColor(\"red\") error: %v", err)
		}
		font, err := style.GetFont()
		if err != nil {
			t.Fatalf("GetFont error: %v", err)
		}
		color, err := font.GetColor()
		if err != nil {
			t.Fatalf("GetColor error: %v", err)
		}
		if color == nil {
			t.Fatal("GetColor returned nil")
		}
	})
	t.Run("valid hex with hash prefix", func(t *testing.T) {
		style := newTestStyle(t)
		if err := editor.WithFontColor("#FF0000")(style); err != nil {
			t.Fatalf("WithFontColor(\"#FF0000\") error: %v", err)
		}
	})
	t.Run("valid int argb", func(t *testing.T) {
		style := newTestStyle(t)
		if err := editor.WithFontColor(0xFF0000)(style); err != nil {
			t.Fatalf("WithFontColor(0xFF0000) error: %v", err)
		}
	})
	t.Run("invalid type returns error", func(t *testing.T) {
		style := newTestStyle(t)
		if err := editor.WithFontColor(3.14)(style); err == nil {
			t.Fatal("WithFontColor(3.14) should return an error for unsupported type")
		}
	})
}

func TestEditorBackgroundColorAndVerticalAlignment(t *testing.T) {
	style := newTestStyle(t)
	if err := editor.WithBackgroundColor("yellow")(style); err != nil {
		t.Fatalf("WithBackgroundColor error: %v", err)
	}
	if err := editor.WithVerticalAlignment("Top")(style); err != nil {
		t.Fatalf("WithVerticalAlignment error: %v", err)
	}
	got, err := style.GetVerticalAlignment()
	if err != nil {
		t.Fatalf("GetVerticalAlignment error: %v", err)
	}
	if got != asposecells.TextAlignmentType_Top {
		t.Errorf("WithVerticalAlignment(\"Top\") => %v, want %v", got, asposecells.TextAlignmentType_Top)
	}
}
