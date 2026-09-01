package tests

import (
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	_ "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/register"
)

// TestCSVRegistered verifies the blank import of the csv saveoptions package
// registers the "csv" format. This guards against a regression of the
// "csv format not registered" bug.
func TestCSVRegistered(t *testing.T) {
	opt := formats.Get("csv")
	if opt == nil {
		t.Fatal("csv format is not registered")
	}
	if got := opt.GetFormat(); got != "csv" {
		t.Errorf("GetFormat() = %q, want %q", got, "csv")
	}
}

// TestAllRegisteredExtensions verifies every saveoptions package blank-imported
// by the register package is actually registered.
func TestAllRegisteredExtensions(t *testing.T) {
	want := []string{
		"bmp", "csv", "dbf", "dif", "docx", "emf", "epub", "gif", "html",
		"jpeg", "jpg", "json", "md", "ods", "pcl", "pdf", "png", "pptx",
		"sql", "svg", "tif", "tiff", "tsv", "txt", "xls", "xlsb", "xlsm",
		"xlsx", "xltm", "xltx", "xml", "xps",
	}
	for _, ext := range want {
		if opt := formats.Get(ext); opt == nil {
			t.Errorf("format %q is not registered", ext)
		}
	}
}
