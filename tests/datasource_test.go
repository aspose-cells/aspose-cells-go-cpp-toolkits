package tests

import (
	"archive/zip"
	"bytes"
	"io"
	"os"
	"path/filepath"
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
)

func TestBytesSource(t *testing.T) {
	payload := []byte("hello bytes")
	src := datasource.BytesSource(payload)

	reader, err := src.Open()
	if err != nil {
		t.Fatalf("Open() error: %v", err)
	}
	data, err := io.ReadAll(reader)
	reader.Close()
	if err != nil {
		t.Fatalf("read error: %v", err)
	}
	if !bytes.Equal(data, payload) {
		t.Errorf("Open() data = %q, want %q", data, payload)
	}
}

func TestFilePathSource(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "input.bin")
	payload := []byte("file data")
	if err := os.WriteFile(path, payload, 0644); err != nil {
		t.Fatalf("WriteFile error: %v", err)
	}

	src := datasource.FilePathSource(path)

	reader, err := src.Open()
	if err != nil {
		t.Fatalf("Open() error: %v", err)
	}
	data, err := io.ReadAll(reader)
	reader.Close()
	if err != nil {
		t.Fatalf("read error: %v", err)
	}
	if !bytes.Equal(data, payload) {
		t.Errorf("Open() data = %q, want %q", data, payload)
	}
}

func TestReaderSource(t *testing.T) {
	payload := []byte("streamed bytes")
	src := datasource.NewReaderSource(bytes.NewReader(payload))

	// First access drains the wrapped stream.
	reader, err := src.Open()
	if err != nil {
		t.Fatalf("Open() error: %v", err)
	}
	data, err := io.ReadAll(reader)
	reader.Close()
	if err != nil {
		t.Fatalf("read error: %v", err)
	}
	if !bytes.Equal(data, payload) {
		t.Errorf("Open() data = %q, want %q", data, payload)
	}
	// Subsequent access must serve the buffered bytes instead of nil.
	reader, err = src.Open()
	if err != nil {
		t.Fatalf("Open() (2nd) error: %v", err)
	}
	data, err = io.ReadAll(reader)
	reader.Close()
	if err != nil {
		t.Fatalf("read error (2nd): %v", err)
	}
	if !bytes.Equal(data, payload) {
		t.Errorf("Open() (2nd) data = %q, want %q", data, payload)
	}
}

func TestFilePathSourceMissingFile(t *testing.T) {
	src := datasource.FilePathSource(filepath.Join(t.TempDir(), "missing.bin"))
	if _, err := src.Open(); err == nil {
		t.Fatal("Open() should return an error for a missing file")
	}
}

func TestBytesSink(t *testing.T) {
	sink := &datasource.BytesSink{}
	if err := sink.Write("", []byte("written")); err != nil {
		t.Fatalf("Write() error: %v", err)
	}
	if got := string(sink.Bytes()); got != "written" {
		t.Errorf("Bytes() = %q, want %q", got, "written")
	}
}

func TestBytesSinkAccumulates(t *testing.T) {
	sink := &datasource.BytesSink{}
	if err := sink.Write("", []byte("first")); err != nil {
		t.Fatalf("Write() error: %v", err)
	}
	if err := sink.Write("", []byte("second")); err != nil {
		t.Fatalf("Write() (2nd) error: %v", err)
	}
	if got := string(sink.Bytes()); got != "firstsecond" {
		t.Errorf("Bytes() = %q, want %q", got, "firstsecond")
	}
}

func TestWriterSink(t *testing.T) {
	var buf bytes.Buffer
	sink := datasource.NewWriterSink(&buf)
	if err := sink.Write("", []byte("sink content")); err != nil {
		t.Fatalf("Write() error: %v", err)
	}
	if buf.String() != "sink content" {
		t.Errorf("writer content = %q, want %q", buf.String(), "sink content")
	}
}

func TestFilePathSink(t *testing.T) {
	path := filepath.Join(t.TempDir(), "out.txt")
	sink := datasource.FilePathSink(path)
	if err := sink.Write("", []byte("sink content")); err != nil {
		t.Fatalf("Write() error: %v", err)
	}

	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatalf("ReadFile error: %v", err)
	}
	if string(data) != "sink content" {
		t.Errorf("file content = %q, want %q", data, "sink content")
	}
}

func TestFolderSink(t *testing.T) {
	dir := filepath.Join(t.TempDir(), "nested", "dir")
	sink := datasource.FolderSink(dir)
	if err := sink.Write("part.xlsx", []byte("part bytes")); err != nil {
		t.Fatalf("Write() error: %v", err)
	}

	data, err := os.ReadFile(filepath.Join(dir, "part.xlsx"))
	if err != nil {
		t.Fatalf("ReadFile error: %v", err)
	}
	if string(data) != "part bytes" {
		t.Errorf("file content = %q, want %q", data, "part bytes")
	}
}

func TestZipSink(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	sink := datasource.NewZipSink(zw)
	if err := sink.Write("a.txt", []byte("content a")); err != nil {
		t.Fatalf("Write() error: %v", err)
	}
	if err := sink.Write("b.txt", []byte("content b")); err != nil {
		t.Fatalf("Write() (2nd) error: %v", err)
	}
	if err := zw.Close(); err != nil {
		t.Fatalf("zip close error: %v", err)
	}

	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("zip reader error: %v", err)
	}
	if len(zr.File) != 2 {
		t.Fatalf("archive entries = %d, want 2", len(zr.File))
	}
	want := map[string]string{"a.txt": "content a", "b.txt": "content b"}
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			t.Fatalf("open entry %s: %v", f.Name, err)
		}
		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			t.Fatalf("read entry %s: %v", f.Name, err)
		}
		if string(data) != want[f.Name] {
			t.Errorf("entry %s content = %q, want %q", f.Name, data, want[f.Name])
		}
	}
}
