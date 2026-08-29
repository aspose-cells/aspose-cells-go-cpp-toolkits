package tests

import (
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

	if !bytes.Equal(src.ByteData(), payload) {
		t.Errorf("ByteData() = %q, want %q", src.ByteData(), payload)
	}

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
	if !bytes.Equal(src.ByteData(), payload) {
		t.Errorf("ByteData() = %q, want %q", src.ByteData(), payload)
	}
}

func TestReaderSource(t *testing.T) {
	payload := []byte("streamed bytes")
	src := datasource.NewReaderSource(io.NopCloser(bytes.NewReader(payload)))

	// First access drains the wrapped stream.
	if !bytes.Equal(src.ByteData(), payload) {
		t.Errorf("ByteData() = %q, want %q", src.ByteData(), payload)
	}
	// Second access must serve the buffered bytes instead of nil.
	if !bytes.Equal(src.ByteData(), payload) {
		t.Errorf("ByteData() (2nd) = %q, want %q", src.ByteData(), payload)
	}
	// Open after buffering still works and is repeatable.
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

func TestFilePathSourceMissingFile(t *testing.T) {
	src := datasource.FilePathSource(filepath.Join(t.TempDir(), "missing.bin"))
	if _, err := src.Open(); err == nil {
		t.Fatal("Open() should return an error for a missing file")
	}
	if got := src.ByteData(); got != nil {
		t.Errorf("ByteData() = %q, want nil for missing file", got)
	}
}

func TestBytesSink(t *testing.T) {
	sink := &datasource.BytesSink{}
	writer, err := sink.Write()
	if err != nil {
		t.Fatalf("Write() error: %v", err)
	}
	if _, err := writer.Write([]byte("written")); err != nil {
		t.Fatalf("Write error: %v", err)
	}
	if err := writer.Close(); err != nil {
		t.Fatalf("Close error: %v", err)
	}
	if got := string(sink.Bytes()); got != "written" {
		t.Errorf("Bytes() = %q, want %q", got, "written")
	}
}

func TestBytesSinkResets(t *testing.T) {
	sink := &datasource.BytesSink{}
	first, _ := sink.Write()
	first.Write([]byte("first"))
	first.Close()

	second, _ := sink.Write()
	second.Write([]byte("second"))
	second.Close()

	if got := string(sink.Bytes()); got != "second" {
		t.Errorf("Bytes() = %q, want %q after reset", got, "second")
	}
}

func TestFilePathSink(t *testing.T) {
	path := filepath.Join(t.TempDir(), "out.txt")
	writer, err := datasource.FilePathSink(path).Write()
	if err != nil {
		t.Fatalf("Write() error: %v", err)
	}
	if _, err := writer.Write([]byte("sink content")); err != nil {
		t.Fatalf("Write error: %v", err)
	}
	writer.Close()

	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatalf("ReadFile error: %v", err)
	}
	if string(data) != "sink content" {
		t.Errorf("file content = %q, want %q", data, "sink content")
	}
}
