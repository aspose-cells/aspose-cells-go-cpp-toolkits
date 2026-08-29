package tests

import (
	"testing"

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
	}{
		{"xlsx", asposecells.FileFormatType_Xlsx, asposecells.SaveFormat_Xlsx},
		{"xlsm", asposecells.FileFormatType_Xlsm, asposecells.SaveFormat_Xlsm},
		{"xlsb", asposecells.FileFormatType_Xlsb, asposecells.SaveFormat_Xlsb},
		{"xls97", asposecells.FileFormatType_Excel97To2003, asposecells.SaveFormat_Excel97To2003},
		{"csv", asposecells.FileFormatType_Csv, asposecells.SaveFormat_Csv},
		{"tsv", asposecells.FileFormatType_Tsv, asposecells.SaveFormat_Tsv},
		{"pdf", asposecells.FileFormatType_Pdf, asposecells.SaveFormat_Pdf},
		{"json", asposecells.FileFormatType_Json, asposecells.SaveFormat_Json},
		{"markdown", asposecells.FileFormatType_Markdown, asposecells.SaveFormat_Markdown},
		{"html", asposecells.FileFormatType_Html, asposecells.SaveFormat_Html},
		{"ods", asposecells.FileFormatType_Ods, asposecells.SaveFormat_Ods},
		{"xps", asposecells.FileFormatType_Xps, asposecells.SaveFormat_Xps},
		{"svg", asposecells.FileFormatType_Svg, asposecells.SaveFormat_Svg},
		{"tiff", asposecells.FileFormatType_Tiff, asposecells.SaveFormat_Tiff},
		{"png", asposecells.FileFormatType_Png, asposecells.SaveFormat_Png},
		{"dbf", asposecells.FileFormatType_Dbf, asposecells.SaveFormat_Dbf},
		{"dif", asposecells.FileFormatType_Dif, asposecells.SaveFormat_Dif},
		{"sql", asposecells.FileFormatType_SqlScript, asposecells.SaveFormat_SqlScript},
		{"docx-group", asposecells.FileFormatType_Docx, asposecells.SaveFormat_Docx},
		{"docm-group", asposecells.FileFormatType_Docm, asposecells.SaveFormat_Docx},
		{"doc-group", asposecells.FileFormatType_Doc, asposecells.SaveFormat_Docx},
		{"dotm-group", asposecells.FileFormatType_Dotm, asposecells.SaveFormat_Docx},
		{"rtf-group", asposecells.FileFormatType_Rtf, asposecells.SaveFormat_Docx},
		{"pptx-group", asposecells.FileFormatType_Pptx, asposecells.SaveFormat_Pptx},
		{"ppt-group", asposecells.FileFormatType_Ppt, asposecells.SaveFormat_Pptx},
		{"ppsm-group", asposecells.FileFormatType_Ppsm, asposecells.SaveFormat_Pptx},
		{"numbers35", asposecells.FileFormatType_Numbers35, asposecells.SaveFormat_Numbers},
		{"unmapped", asposecells.FileFormatType_Unknown, asposecells.SaveFormat_Auto},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			got := formats.FileFormatToSaveFormat(tc.formatType)
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
