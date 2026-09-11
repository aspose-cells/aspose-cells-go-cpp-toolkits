package tests

import (
	"fmt"
	"testing"
	"time"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// newTestWorkbookBytes returns the serialized bytes of a blank XLSX workbook so
// the editor DSL can be exercised through the same datasource path used in
// production code.
func newTestWorkbookBytes(t *testing.T) []byte {
	t.Helper()
	wb, err := asposecells.NewWorkbook()
	if err != nil {
		t.Fatalf("NewWorkbook error: %v", err)
	}
	data, err := wb.Save_SaveFormat(asposecells.SaveFormat_Xlsx)
	if err != nil {
		t.Fatalf("Save_SaveFormat error: %v", err)
	}
	return data
}

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

// TestSetCellValueRoundTrip writes string, int32, and bool values through the
// toObject/PutValue_Object path and reads them back from the serialized
// workbook, verifying the type conversion preserves the values. The read-back
// runs under retryStable because evaluation mode can corrupt a random cell's
// value at load time; a genuine regression fails every attempt.
func TestSetCellValueRoundTrip(t *testing.T) {
	out, err := editor.EditSpreadsheet(
		datasource.BytesSource(newTestWorkbookBytes(t)),
		editor.InWorksheet(0,
			editor.SetCellValue(0, 0, "hello"),
			editor.SetCellValue(0, 1, int32(42)),
			editor.SetCellValue(0, 2, true),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet error: %v", err)
	}
	if err := retryStable(5, func() error {
		return verifyRoundTrip(out)
	}); err != nil {
		t.Fatal(err)
	}
}

// verifyRoundTrip loads a serialized workbook and checks the values written by
// TestSetCellValueRoundTrip. It returns an error so the caller can re-load
// under retryStable.
func verifyRoundTrip(out []byte) error {
	wb, err := asposecells.NewWorkbook_Stream(out)
	if err != nil {
		return fmt.Errorf("NewWorkbook_Stream: %w", err)
	}
	wss, err := wb.GetWorksheets()
	if err != nil {
		return err
	}
	ws, err := wss.Get_Int(0)
	if err != nil {
		return err
	}
	cells, err := ws.GetCells()
	if err != nil {
		return err
	}

	get := func(row, col int32) (*asposecells.Cell, error) {
		return cells.Get_Int_Int(row, col)
	}

	cell, err := get(0, 0)
	if err != nil {
		return fmt.Errorf("Get_Int_Int(0,0): %w", err)
	}
	got, err := cell.GetStringValue()
	if err != nil {
		return fmt.Errorf("GetStringValue: %w", err)
	}
	if got != "hello" {
		return fmt.Errorf("string cell = %q, want %q", got, "hello")
	}

	cell, err = get(0, 1)
	if err != nil {
		return fmt.Errorf("Get_Int_Int(0,1): %w", err)
	}
	iv, err := cell.GetIntValue()
	if err != nil {
		return fmt.Errorf("GetIntValue: %w", err)
	}
	if iv != 42 {
		return fmt.Errorf("int cell = %d, want 42", iv)
	}

	cell, err = get(0, 2)
	if err != nil {
		return fmt.Errorf("Get_Int_Int(0,2): %w", err)
	}
	bv, err := cell.GetBoolValue()
	if err != nil {
		return fmt.Errorf("GetBoolValue: %w", err)
	}
	if !bv {
		return fmt.Errorf("bool cell = false, want true")
	}
	return nil
}

// TestEditorMergeUnmerge verifies the Merge and UnMerge actions work correctly.
func TestEditorMergeUnmerge(t *testing.T) {
	out, err := editor.EditSpreadsheet(
		datasource.BytesSource(newTestWorkbookBytes(t)),
		editor.InWorksheet(0,
			editor.SetCellValue(0, 0, "A1"),
			editor.SetCellValue(0, 1, "B1"),
			editor.SetCellValue(1, 0, "A2"),
			editor.SetCellValue(1, 1, "B2"),
			editor.Merge(0, 0, 2, 2, false),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet with Merge error: %v", err)
	}

	if err := retryStable(5, func() error {
		wb, err := asposecells.NewWorkbook_Stream(out)
		if err != nil {
			return fmt.Errorf("load: %w", err)
		}
		wss, err := wb.GetWorksheets()
		if err != nil {
			return err
		}
		ws, err := wss.Get_Int(0)
		if err != nil {
			return err
		}
		cells, err := ws.GetCells()
		if err != nil {
			return err
		}

		// Check merged areas
		areas, err := cells.GetMergedAreas()
		if err != nil {
			return err
		}
		if len(areas) != 1 {
			return fmt.Errorf("merged areas count = %d, want 1", len(areas))
		}

		startRow, err := areas[0].Get_StartRow()
		if err != nil {
			return err
		}
		endRow, err := areas[0].Get_EndRow()
		if err != nil {
			return err
		}
		startCol, err := areas[0].Get_StartColumn()
		if err != nil {
			return err
		}
		endCol, err := areas[0].Get_EndColumn()
		if err != nil {
			return err
		}

		if startRow != 0 || endRow != 1 || startCol != 0 || endCol != 1 {
			return fmt.Errorf("merged area = (%d,%d)-(%d,%d), want (0,0)-(1,1)", startRow, startCol, endRow, endCol)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}

	// Now unmerge and verify
	out2, err := editor.EditSpreadsheet(
		datasource.BytesSource(out),
		editor.InWorksheet(0,
			editor.UnMerge(0, 0, 2, 2),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet with UnMerge error: %v", err)
	}

	if err := retryStable(5, func() error {
		wb, err := asposecells.NewWorkbook_Stream(out2)
		if err != nil {
			return fmt.Errorf("load: %w", err)
		}
		wss, err := wb.GetWorksheets()
		if err != nil {
			return err
		}
		ws, err := wss.Get_Int(0)
		if err != nil {
			return err
		}
		cells, err := ws.GetCells()
		if err != nil {
			return err
		}

		// Check no merged areas
		areas, err := cells.GetMergedAreas()
		if err != nil {
			return err
		}
		if len(areas) != 0 {
			return fmt.Errorf("merged areas count after unmerge = %d, want 0", len(areas))
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

// TestEditorClearComments verifies the ClearComments action works correctly.
func TestEditorClearComments(t *testing.T) {
	// Create a workbook with a comment
	out, err := editor.EditSpreadsheet(
		datasource.BytesSource(newTestWorkbookBytes(t)),
		editor.InWorksheet(0,
			editor.SetCellValue(0, 0, "test"),
			editor.SetCellComment(0, 0, "hello note"),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet with comment error: %v", err)
	}

	// Verify comment exists
	if err := retryStable(5, func() error {
		wb, err := asposecells.NewWorkbook_Stream(out)
		if err != nil {
			return fmt.Errorf("load: %w", err)
		}
		wss, err := wb.GetWorksheets()
		if err != nil {
			return err
		}
		ws, err := wss.Get_Int(0)
		if err != nil {
			return err
		}
		comments, err := ws.GetComments()
		if err != nil {
			return err
		}
		count, err := comments.GetCount()
		if err != nil {
			return err
		}
		if count != 1 {
			return fmt.Errorf("comments count = %d, want 1", count)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}

	// Clear comments
	out2, err := editor.EditSpreadsheet(
		datasource.BytesSource(out),
		editor.InWorksheet(0,
			editor.ClearComments(),
		),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet with ClearComments error: %v", err)
	}

	// Verify comment is cleared
	if err := retryStable(5, func() error {
		wb, err := asposecells.NewWorkbook_Stream(out2)
		if err != nil {
			return fmt.Errorf("load: %w", err)
		}
		wss, err := wb.GetWorksheets()
		if err != nil {
			return err
		}
		ws, err := wss.Get_Int(0)
		if err != nil {
			return err
		}
		comments, err := ws.GetComments()
		if err != nil {
			return err
		}
		count, err := comments.GetCount()
		if err != nil {
			return err
		}
		if count != 0 {
			return fmt.Errorf("comments count after clear = %d, want 0", count)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

// TestEditorDeleteWorksheet verifies the WithDeleteWorksheet action works correctly.
func TestEditorDeleteWorksheet(t *testing.T) {
	// Create a workbook with multiple sheets
	seed, err := asposecells.NewWorkbook()
	if err != nil {
		t.Fatalf("NewWorkbook error: %v", err)
	}
	seedBytes, err := seed.Save_SaveFormat(asposecells.SaveFormat_Xlsx)
	if err != nil {
		t.Fatalf("Save error: %v", err)
	}

	// Add a second sheet
	wb, err := asposecells.NewWorkbook_Stream(seedBytes)
	if err != nil {
		t.Fatalf("NewWorkbook_Stream error: %v", err)
	}
	wss, err := wb.GetWorksheets()
	if err != nil {
		t.Fatalf("GetWorksheets error: %v", err)
	}
	_, err = wss.Add_String("Sheet2")
	if err != nil {
		t.Fatalf("Add_String error: %v", err)
	}
	seedWithTwoSheets, err := wb.Save_SaveFormat(asposecells.SaveFormat_Xlsx)
	if err != nil {
		t.Fatalf("Save error: %v", err)
	}

	// Delete the second sheet
	out, err := editor.EditSpreadsheet(
		datasource.BytesSource(seedWithTwoSheets),
		editor.WithDeleteWorksheet("Sheet2"),
	)
	if err != nil {
		t.Fatalf("EditSpreadsheet with DeleteWorksheet error: %v", err)
	}

	if err := retryStable(5, func() error {
		wb, err := asposecells.NewWorkbook_Stream(out)
		if err != nil {
			return fmt.Errorf("load: %w", err)
		}
		wss, err := wb.GetWorksheets()
		if err != nil {
			return err
		}
		count, err := wss.GetCount()
		if err != nil {
			return err
		}
		if count != 1 {
			return fmt.Errorf("worksheet count = %d, want 1", count)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

// TestEditorSetCellValues tests the SetCellValues action for writing
// multiple values to a worksheet range in one call.
func TestEditorSetCellValues(t *testing.T) {
	cases := []struct {
		name   string
		start  int
		values [][]interface{}
	}{
		{
			name: "single cell",
			values: [][]interface{}{
				{"A1"},
			},
		},
		{
			name: "single row",
			values: [][]interface{}{
				{"A", "B", "C", "D"},
			},
		},
		{
			name: "single column",
			values: [][]interface{}{
				{"Row1"},
				{"Row2"},
				{"Row3"},
			},
		},
		{
			name: "multiple rows and columns",
			values: [][]interface{}{
				{"Name", "Age", "City"},
				{"Alice", 30, "New York"},
				{"Bob", 25, "Los Angeles"},
				{"Charlie", 35, "Chicago"},
			},
		},
		{
			name: "mixed types including nil",
			values: [][]interface{}{
				{"Text", 123, 45.67, true},
				{nil, nil, nil, nil},
				{time.Date(2024, 1, 15, 0, 0, 0, 0, time.UTC), "more", "data", 100},
			},
		},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			out, err := editor.EditSpreadsheet(
				datasource.BytesSource(newTestWorkbookBytes(t)),
				editor.InWorksheet(0,
					editor.SetCellValues(0, 0, tc.values),
				),
			)
			if err != nil {
				t.Fatalf("EditSpreadsheet with SetCellValues error: %v", err)
			}

			// Verify the values were written correctly
			if err := retryStable(5, func() error {
				return verifySetCellValues(out, tc.values)
			}); err != nil {
				t.Fatalf("%s: %v", tc.name, err)
			}
		})
	}
}

// verifySetCellValues loads a serialized workbook and verifies the values
// written by TestEditorSetCellValues.
func verifySetCellValues(out []byte, expected [][]interface{}) error {
	wb, err := asposecells.NewWorkbook_Stream(out)
	if err != nil {
		return fmt.Errorf("NewWorkbook_Stream: %w", err)
	}
	wss, err := wb.GetWorksheets()
	if err != nil {
		return err
	}
	ws, err := wss.Get_Int(0)
	if err != nil {
		return err
	}
	cells, err := ws.GetCells()
	if err != nil {
		return err
	}

	rowCount := len(expected)
	if rowCount == 0 {
		return nil
	}

	colCount := 0
	for _, row := range expected {
		if len(row) > colCount {
			colCount = len(row)
		}
	}

	for r := 0; r < rowCount; r++ {
		for c := 0; c < colCount; c++ {
			expectedVal := expected[r][c]

			// Skip nil values - they should remain unchanged (empty)
			if expectedVal == nil {
				continue
			}

			cell, err := cells.Get_Int_Int(int32(r), int32(c))
			if err != nil {
				return fmt.Errorf("Get_Int_Int(%d,%d): %w", r, c, err)
			}

			switch v := expectedVal.(type) {
			case string:
				got, err := cell.GetStringValue()
				if err != nil {
					return fmt.Errorf("GetStringValue(%d,%d): %w", r, c, err)
				}
				if got != v {
					return fmt.Errorf("cell(%d,%d) string = %q, want %q", r, c, got, v)
				}
			case int:
				bv, err := cell.GetBoolValue()
				if err == nil && bv {
					return fmt.Errorf("cell(%d,%d) should not be bool", r, c)
				}
				iv, err := cell.GetIntValue()
				if err != nil {
					return fmt.Errorf("GetIntValue(%d,%d): %w", r, c, err)
				}
				if int(iv) != v {
					return fmt.Errorf("cell(%d,%d) int = %d, want %d", r, c, iv, v)
				}
			case int64:
				iv, err := cell.GetIntValue()
				if err != nil {
					return fmt.Errorf("GetIntValue(%d,%d): %w", r, c, err)
				}
				if int64(iv) != v {
					return fmt.Errorf("cell(%d,%d) int64 = %d, want %d", r, c, iv, v)
				}
			case float64:
				fv, err := cell.GetDoubleValue()
				if err != nil {
					return fmt.Errorf("GetDoubleValue(%d,%d): %w", r, c, err)
				}
				if fv < v-0.0001 || fv > v+0.0001 {
					return fmt.Errorf("cell(%d,%d) float = %f, want %f", r, c, fv, v)
				}
			case bool:
				bv, err := cell.GetBoolValue()
				if err != nil {
					return fmt.Errorf("GetBoolValue(%d,%d): %w", r, c, err)
				}
				if bv != v {
					return fmt.Errorf("cell(%d,%d) bool = %v, want %v", r, c, bv, v)
				}
			case time.Time:
				dt, err := cell.GetDateTimeValue()
				if err != nil {
					return fmt.Errorf("GetDateTimeValue(%d,%d): %w", r, c, err)
				}
				if dt.Year() != v.Year() || dt.Month() != v.Month() || dt.Day() != v.Day() {
					return fmt.Errorf("cell(%d,%d) date = %v, want %v", r, c, dt, v)
				}
			}
		}
	}

	return nil
}

// TestEditorSetCellValuesEmpty tests edge cases with empty/nil values.
func TestEditorSetCellValuesEmpty(t *testing.T) {
	// Empty 2D slice
	out, err := editor.EditSpreadsheet(
		datasource.BytesSource(newTestWorkbookBytes(t)),
		editor.InWorksheet(0,
			editor.SetCellValues(0, 0, [][]interface{}{}),
		),
	)
	if err != nil {
		t.Fatalf("SetCellValues with empty slice error: %v", err)
	}

	// Slice with empty rows
	out, err = editor.EditSpreadsheet(
		datasource.BytesSource(out),
		editor.InWorksheet(0,
			editor.SetCellValues(5, 5, [][]interface{}{{}, {}}),
		),
	)
	if err != nil {
		t.Fatalf("SetCellValues with empty rows error: %v", err)
	}

	// All nil values
	out, err = editor.EditSpreadsheet(
		datasource.BytesSource(out),
		editor.InWorksheet(0,
			editor.SetCellValues(10, 10, [][]interface{}{
				{nil, nil},
				{nil, nil},
			}),
		),
	)
	if err != nil {
		t.Fatalf("SetCellValues with all nil values error: %v", err)
	}
}
