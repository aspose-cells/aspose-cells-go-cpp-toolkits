package tests

import (
	"errors"
	"testing"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	engine "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/engine"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// TestFileFormatToSaveFormat exercises the input-format to save-format mapping
// that used to be exported as formats.FileFormatToSaveFormat. It is internal
// now: the mapping names engine types on both sides, so exposing it would tie
// callers of the public formats package to the binding (docs/design.md §11).
func TestFileFormatToSaveFormat(t *testing.T) {
	cases := []struct {
		name       string
		formatType asposecells.FileFormatType
		want       asposecells.SaveFormat
		wantErr    bool
	}{
		{"xlsx", asposecells.FileFormatType_Xlsx, asposecells.SaveFormat_Xlsx, false},
		{"xlsm", asposecells.FileFormatType_Xlsm, asposecells.SaveFormat_Xlsm, false},
		{"xlsb", asposecells.FileFormatType_Xlsb, asposecells.SaveFormat_Xlsb, false},
		{"xls97", asposecells.FileFormatType_Excel97To2003, asposecells.SaveFormat_Excel97To2003, false},
		{"csv", asposecells.FileFormatType_Csv, asposecells.SaveFormat_Csv, false},
		{"tsv", asposecells.FileFormatType_Tsv, asposecells.SaveFormat_Tsv, false},
		{"pdf", asposecells.FileFormatType_Pdf, asposecells.SaveFormat_Pdf, false},
		{"json", asposecells.FileFormatType_Json, asposecells.SaveFormat_Json, false},
		{"markdown", asposecells.FileFormatType_Markdown, asposecells.SaveFormat_Markdown, false},
		{"html", asposecells.FileFormatType_Html, asposecells.SaveFormat_Html, false},
		{"ods", asposecells.FileFormatType_Ods, asposecells.SaveFormat_Ods, false},
		{"xps", asposecells.FileFormatType_Xps, asposecells.SaveFormat_Xps, false},
		{"svg", asposecells.FileFormatType_Svg, asposecells.SaveFormat_Svg, false},
		{"tiff", asposecells.FileFormatType_Tiff, asposecells.SaveFormat_Tiff, false},
		{"png", asposecells.FileFormatType_Png, asposecells.SaveFormat_Png, false},
		{"dbf", asposecells.FileFormatType_Dbf, asposecells.SaveFormat_Dbf, false},
		{"dif", asposecells.FileFormatType_Dif, asposecells.SaveFormat_Dif, false},
		{"sql", asposecells.FileFormatType_SqlScript, asposecells.SaveFormat_SqlScript, false},
		{"docx-group", asposecells.FileFormatType_Docx, asposecells.SaveFormat_Docx, false},
		{"docm-group", asposecells.FileFormatType_Docm, asposecells.SaveFormat_Docx, false},
		{"doc-group", asposecells.FileFormatType_Doc, asposecells.SaveFormat_Docx, false},
		{"dotm-group", asposecells.FileFormatType_Dotm, asposecells.SaveFormat_Docx, false},
		{"rtf-group", asposecells.FileFormatType_Rtf, asposecells.SaveFormat_Docx, false},
		{"pptx-group", asposecells.FileFormatType_Pptx, asposecells.SaveFormat_Pptx, false},
		{"ppt-group", asposecells.FileFormatType_Ppt, asposecells.SaveFormat_Pptx, false},
		{"ppsm-group", asposecells.FileFormatType_Ppsm, asposecells.SaveFormat_Pptx, false},
		{"numbers35", asposecells.FileFormatType_Numbers35, asposecells.SaveFormat_Numbers, false},
		// An unmapped format is an error, not SaveFormat_Auto: Auto would let
		// the engine pick a format, quietly saving a file as something other
		// than what was asked for.
		{"unmapped", asposecells.FileFormatType_Unknown, 0, true},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			got, err := engine.FileFormatToSaveFormat(tc.formatType)
			if tc.wantErr {
				if !errors.Is(err, toolkiterrors.ErrUnsupportedFormat) {
					t.Fatalf("FileFormatToSaveFormat(%v) error = %v, want ErrUnsupportedFormat", tc.formatType, err)
				}
				return
			}
			if err != nil {
				t.Fatalf("FileFormatToSaveFormat(%v) unexpected error: %v", tc.formatType, err)
			}
			if got != tc.want {
				t.Errorf("FileFormatToSaveFormat(%v) = %v, want %v", tc.formatType, got, tc.want)
			}
		})
	}
}
