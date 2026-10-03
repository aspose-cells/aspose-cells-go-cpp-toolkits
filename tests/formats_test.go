package tests

import (
	"errors"
	"testing"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// stubOption is a minimal SaveOption used to exercise the formats registry
// without depending on any saveoptions implementation.
type stubOption struct{}

func (stubOption) Apply(source []byte) ([]byte, error) { return source, nil }
func (stubOption) GetFormat() string                   { return "stub" }

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
			got, err := formats.FileFormatToSaveFormat(tc.formatType)
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

// The fake extensions below must not collide with real registered formats,
// so a distinctive "zzz-" prefix is used.
func TestRegistryCaseInsensitive(t *testing.T) {
	formats.Register("zzz-test-ext", func() saveoptions.SaveOption { return stubOption{} })
	defer formats.Unregister("zzz-test-ext")

	for _, key := range []string{"zzz-test-ext", "ZZZ-TEST-EXT", "ZzZ-tEsT-eXt"} {
		opt := formats.Get(key)
		if opt == nil {
			t.Fatalf("Get(%q) returned nil", key)
		}
		if got := opt.GetFormat(); got != "stub" {
			t.Errorf("Get(%q).GetFormat() = %q, want %q", key, got, "stub")
		}
	}
}

func TestRegistryGetUnknownReturnsNil(t *testing.T) {
	if opt := formats.Get("definitely-not-a-format"); opt != nil {
		t.Errorf("Get(unknown) = %v, want nil", opt)
	}
}

func TestListContainsRegistered(t *testing.T) {
	formats.Register("zzz-list-check", func() saveoptions.SaveOption { return stubOption{} })
	defer formats.Unregister("zzz-list-check")
	found := false
	for _, ext := range formats.List() {
		if ext == "zzz-list-check" {
			found = true
			break
		}
	}
	if !found {
		t.Errorf("List() = %v, want it to contain zzz-list-check", formats.List())
	}
}

// TestUnregisterRemovesExtension verifies Unregister cleans the registry, which
// is how tests avoid polluting the process-wide registry with fake extensions.
func TestUnregisterRemovesExtension(t *testing.T) {
	formats.Register("zzz-unreg", func() saveoptions.SaveOption { return stubOption{} })
	if formats.Get("zzz-unreg") == nil {
		t.Fatal("zzz-unreg should be registered before Unregister")
	}
	formats.Unregister("zzz-unreg")
	if formats.Get("zzz-unreg") != nil {
		t.Error("zzz-unreg should return nil after Unregister")
	}
}

func TestListSorted(t *testing.T) {
	exts := formats.List()
	for i := 1; i < len(exts); i++ {
		if exts[i-1] > exts[i] {
			t.Fatalf("List() is not sorted: %q > %q at index %d", exts[i-1], exts[i], i)
		}
	}
}
