// Package datasource abstracts spreadsheet input and output.
//
// DataSource implementations provide readable input for processing; DataSink
// implementations accept output written by the toolkit. The converter, editor,
// manipulator, and transfer packages take these abstractions instead of
// concrete file/bytes/stream types, so one signature covers every input and
// output shape and the number of public entry points stays small.
package datasource

import (
	"archive/zip"
	"bytes"
	"io"
	"os"
	"path/filepath"
	"sync"
)

// DataSource abstracts a readable input. Any type that implements Open can
// serve as a data provider; Open returns an io.ReadCloser that the caller
// (the toolkit, via internal/aspose/cells.ReadSource) drains and closes.
type DataSource interface {
	Open() (io.ReadCloser, error)
}

// FilePathSource is a DataSource backed by a local file path.
type FilePathSource string

// Open opens the file at the specified path. It directly calls os.Open, so if
// the file does not exist or lacks permissions, the corresponding system error
// is returned.
func (p FilePathSource) Open() (io.ReadCloser, error) {
	return os.Open(string(p))
}

// BytesSource is a DataSource backed by an in-memory byte slice.
type BytesSource []byte

// Open wraps the byte slice into an io.ReadCloser. Since the data is in
// memory, this operation does not fail.
func (b BytesSource) Open() (io.ReadCloser, error) {
	return io.NopCloser(bytes.NewReader(b)), nil
}

// ReaderSource is an adapter for an already-open io.Reader, such as a file or
// HTTP response body. It lets a stream be used wherever a DataSource is
// expected.
//
// The wrapped stream is consumed lazily on the first access and buffered, so
// Open may be called repeatedly; subsequent calls serve the same buffered
// bytes. The caller remains responsible for closing the underlying stream.
type ReaderSource struct {
	mu     sync.Mutex
	reader io.Reader
	buf    []byte
}

// NewReaderSource wraps an already-open stream as a DataSource. The returned
// source buffers the stream on first use.
func NewReaderSource(r io.Reader) *ReaderSource {
	return &ReaderSource{reader: r}
}

// Open returns an io.ReadCloser over the source contents. Because the source
// is buffered on first access, every call returns a fresh reader over the
// same bytes.
func (r *ReaderSource) Open() (io.ReadCloser, error) {
	data, err := r.readAll()
	if err != nil {
		return nil, err
	}
	return io.NopCloser(bytes.NewReader(data)), nil
}

func (r *ReaderSource) readAll() ([]byte, error) {
	r.mu.Lock()
	defer r.mu.Unlock()
	if r.buf != nil {
		return r.buf, nil
	}
	if r.reader == nil {
		return nil, nil
	}
	data, err := io.ReadAll(r.reader)
	r.reader = nil
	if err != nil {
		return nil, err
	}
	r.buf = data
	return data, nil
}

// DataSink abstracts a writable output. The name parameter is used by
// multi-output operations (for example manipulator.Split writing one file or
// archive entry per worksheet); single-output sinks ignore it.
type DataSink interface {
	Write(name string, data []byte) error
}

// FilePathSink writes each Write call to a file at the given path. name is
// ignored.
type FilePathSink string

// Write writes data to the file at the sink's path, truncating any existing
// file.
func (p FilePathSink) Write(_ string, data []byte) error {
	return os.WriteFile(string(p), data, 0o644)
}

// WriterSink writes each Write call to the wrapped io.Writer. name is ignored.
type WriterSink struct {
	w io.Writer
}

// NewWriterSink wraps an io.Writer as a DataSink.
func NewWriterSink(w io.Writer) *WriterSink {
	return &WriterSink{w: w}
}

// Write forwards data to the wrapped writer.
func (s WriterSink) Write(_ string, data []byte) error {
	_, err := s.w.Write(data)
	return err
}

// BytesSink accumulates all Write calls into an in-memory buffer, so a caller
// that wants the result as bytes can hand a BytesSink to any toolkit entry
// point and read the output back with Bytes.
type BytesSink struct {
	buf bytes.Buffer
}

// Write appends data to the sink's buffer. name is ignored.
func (b *BytesSink) Write(_ string, data []byte) error {
	_, err := b.buf.Write(data)
	return err
}

// Bytes returns the accumulated output.
func (b *BytesSink) Bytes() []byte {
	return b.buf.Bytes()
}

// FolderSink writes each Write call to a file named name inside the given
// folder, creating the folder if needed. It backs manipulator.Split's
// per-worksheet file output.
type FolderSink string

// Write creates the folder if needed and writes data to
// <folder>/<name>.
func (f FolderSink) Write(name string, data []byte) error {
	if err := os.MkdirAll(string(f), 0o755); err != nil {
		return err
	}
	return os.WriteFile(filepath.Join(string(f), name), data, 0o644)
}

// ZipSink writes each Write call as a new entry named name inside a
// zip.Writer, backing manipulator.Split's per-worksheet archive output.
type ZipSink struct {
	zw *zip.Writer
}

// NewZipSink wraps a zip.Writer as a DataSink. The caller closes the zip.Writer
// after the operation completes to finalize the archive.
func NewZipSink(zw *zip.Writer) *ZipSink {
	return &ZipSink{zw: zw}
}

// Write creates an entry named name in the archive and writes data to it.
func (z ZipSink) Write(name string, data []byte) error {
	entry, err := z.zw.Create(name)
	if err != nil {
		return err
	}
	_, err = entry.Write(data)
	return err
}
