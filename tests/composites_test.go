package tests

import (
	"bytes"
	"errors"
	"os"
	"path/filepath"
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/converter"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	_ "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/register"
)

// TestConvertNilSentinels verifies the composite entry points classify nil
// source, sink, and option with errors.Is instead of panicking. The nil checks
// short-circuit before any engine call, so no workbook is needed.
func TestConvertNilSentinels(t *testing.T) {
	var sink datasource.BytesSink
	opt := formats.Get("xlsx")

	t.Run("nil source", func(t *testing.T) {
		err := converter.Convert(nil, opt, &sink)
		if !errors.Is(err, toolkiterrors.ErrDataSourceNil) {
			t.Fatalf("Convert(nil source) error = %v, want ErrDataSourceNil", err)
		}
	})

	t.Run("nil sink", func(t *testing.T) {
		err := converter.Convert(datasource.BytesSource([]byte("x")), opt, nil)
		if !errors.Is(err, toolkiterrors.ErrDataSinkNil) {
			t.Fatalf("Convert(nil sink) error = %v, want ErrDataSinkNil", err)
		}
	})

	t.Run("nil option", func(t *testing.T) {
		err := converter.Convert(datasource.BytesSource([]byte("x")), nil, &sink)
		if !errors.Is(err, toolkiterrors.ErrSaveOptionNil) {
			t.Fatalf("Convert(nil option) error = %v, want ErrSaveOptionNil", err)
		}
	})
}

// TestConvertSinkRouting verifies converter.Convert delivers its output to
// whichever sink the caller chooses: a FilePathSink lands on disk at the given
// path, a WriterSink reaches the wrapped writer, and a BytesSink accumulates
// the bytes. Each Convert call renders the workbook separately, and XLSX output
// is non-deterministic in evaluation mode, so the assertions check destination
// correctness (non-empty, at the right place) rather than byte equality.
func TestConvertSinkRouting(t *testing.T) {
	src := datasource.BytesSource(newTestWorkbookBytes(t))
	opt := formats.Get("xlsx")

	// BytesSink captures the result in memory.
	var bytesSink datasource.BytesSink
	if err := converter.Convert(src, opt, &bytesSink); err != nil {
		t.Fatalf("Convert to BytesSink: %v", err)
	}
	if len(bytesSink.Bytes()) == 0 {
		t.Fatal("Convert produced empty output")
	}

	// FilePathSink must write a non-empty file at the given path.
	path := filepath.Join(t.TempDir(), "out.xlsx")
	if err := converter.Convert(src, opt, datasource.FilePathSink(path)); err != nil {
		t.Fatalf("Convert to FilePathSink: %v", err)
	}
	got, err := os.ReadFile(path)
	if err != nil {
		t.Fatalf("ReadFile: %v", err)
	}
	if len(got) == 0 {
		t.Fatal("FilePathSink produced an empty file")
	}

	// WriterSink must forward the output to the wrapped writer.
	var buf bytes.Buffer
	if err := converter.Convert(src, opt, datasource.NewWriterSink(&buf)); err != nil {
		t.Fatalf("Convert to WriterSink: %v", err)
	}
	if buf.Len() == 0 {
		t.Fatal("WriterSink produced no output")
	}
}

// TestConvertDeprecatedWrappers verifies the retained deprecated wrappers still
// produce usable output, so existing callers keep working.
func TestConvertDeprecatedWrappers(t *testing.T) {
	src := datasource.BytesSource(newTestWorkbookBytes(t))
	opt := formats.Get("xlsx")

	bytesOut, err := converter.ConvertSpreadsheet(src, opt)
	if err != nil {
		t.Fatalf("ConvertSpreadsheet: %v", err)
	}
	if len(bytesOut) == 0 {
		t.Fatal("ConvertSpreadsheet returned empty output")
	}

	path := filepath.Join(t.TempDir(), "copy.xlsx")
	if err := converter.ConvertSpreadsheetToFile(filepath.Join(t.TempDir(), "in.xlsx"), "out.pdf"); err != nil {
		// Expected to fail on the missing input; only assert the deprecated
		// wrapper itself is reachable and surfaces the input error, not a
		// format-classification error.
		if errors.Is(err, toolkiterrors.ErrUnsupportedFormat) {
			t.Fatalf("missing input misclassified as ErrUnsupportedFormat: %v", err)
		}
	}
	_ = path
}
