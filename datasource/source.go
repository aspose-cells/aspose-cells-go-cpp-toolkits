// Package datasource abstracts spreadsheet input and output.
//
// DataSource implementations provide bytes for processing, and DataSink
// implementations accept bytes written by the toolkit. Sources and sinks can
// back onto files, in-memory buffers, or already-open streams, letting the
// converter, editor, manipulator, and transfer packages stay agnostic about
// where data comes from or goes to.
package datasource

import (
	"bytes"
	"io"
	"os"
	"sync"
)

// The DataSource defines the general interface for data sources.
// Any type that implements this interface can serve as a data provider.
// The Open method is responsible for returning an io.ReadCloser, and the caller must close it after
type DataSource interface {
	Open() (io.ReadCloser, error)
	ByteData() []byte
}

// FilePathSource is an implementation of a data source based on local file system paths.
// It implements the DataSource interface and is used to read files from the specified path.
type FilePathSource string

// Open Opens the file at the specified path.
// It directly calls os.Open, so if the file does not exist or does not have the required permissions, it will return the corresponding system error.
func (p FilePathSource) Open() (io.ReadCloser, error) {
	return os.Open(string(p))
}

func (p FilePathSource) ByteData() []byte {
	reader, errOpen := p.Open()
	if errOpen != nil {
		return nil
	}
	data, errRead := io.ReadAll(reader)
	if errRead != nil {
		return nil
	}
	reader.Close()
	return data
}

// BytesSource is an implementation of a data source based on in-memory byte slices.
// This is very useful in unit tests or when processing data that has been loaded into memory.
type BytesSource []byte

// Open wraps the byte slice into an io.ReadCloser.
// Since the data is in memory, this operation usually does not fail (returns nil error).
// io.NopCloser is used to wrap an io.Reader into an io.ReadCloser, and its Close method is an empty operation.
func (b BytesSource) Open() (io.ReadCloser, error) {
	return io.NopCloser(bytes.NewReader(b)), nil
}

func (b BytesSource) ByteData() []byte {
	return b
}

// ReaderSource is an adapter for an already-open io.ReadCloser, such as an
// HTTP response body. It lets an open stream be used wherever a DataSource is
// expected.
//
// The wrapped stream is consumed lazily on the first access and buffered, so
// Open and ByteData may be called repeatedly; subsequent calls serve the same
// buffered bytes.
type ReaderSource struct {
	mu     sync.Mutex
	reader io.ReadCloser
	buf    []byte
}

// NewReaderSource wraps an already-open stream as a DataSource. The returned
// source buffers the stream on first use; the wrapped stream is closed once
// its contents have been read.
func NewReaderSource(r io.ReadCloser) *ReaderSource {
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

// ByteData returns the source contents. It is safe to call repeatedly; the
// wrapped stream is drained and closed exactly once, after which the bytes are
// served from an internal buffer.
func (r *ReaderSource) ByteData() []byte {
	data, _ := r.readAll()
	return data
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
	if closeErr := r.reader.Close(); closeErr != nil && err == nil {
		err = closeErr
	}
	r.reader = nil
	if err != nil {
		return nil, err
	}
	r.buf = data
	return data, nil
}

// DataSink defines the general interface for data destinations.
// Any type that implements this interface can serve as a data writer.
// The Write method returns an io.WriteCloser, and the caller must close it after writing.
type DataSink interface {
	Write() (io.WriteCloser, error)
}

// FilePathSink is an implementation of a data sink based on local file system paths.
// It implements the DataSink interface and is used to write data to the specified file path.
type FilePathSink string

// Write creates or truncates the file at the specified path and returns a WriteCloser.
// It directly calls os.Create, so if the directory does not exist or lacks permissions, it will return an error.
func (p FilePathSink) Write() (io.WriteCloser, error) {
	return os.Create(string(p))
}

// BytesSink is an in-memory implementation of a data sink.
// It is useful for capturing written data into a byte slice, typically used in testing or buffering.
type BytesSink struct {
	buf bytes.Buffer
}

// Write returns a wrapper around the internal buffer.
// The internal buffer is reset before writing to ensure a clean state.
func (b *BytesSink) Write() (io.WriteCloser, error) {
	b.buf.Reset()
	return nopWriteCloser{&b.buf}, nil
}

// Bytes returns the accumulated bytes written to the sink.
func (b *BytesSink) Bytes() []byte {
	return b.buf.Bytes()
}

// nopWriteCloser wraps an io.Writer to satisfy the io.WriteCloser interface.
// Its Close method is a no-op.
type nopWriteCloser struct {
	io.Writer
}

// Close is a no-op for in-memory buffers.
func (nopWriteCloser) Close() error {
	return nil
}

// FileStore is an implementation that supports both reading and writing to a local file.
// It implements both DataSource and DataSink interfaces.
// This is useful when a single entity represents a file that needs to be both read from and written to.
type FileStore string

// Open opens the file at the specified path for reading.
func (f FileStore) Open() (io.ReadCloser, error) {
	return os.Open(string(f))
}

// Write creates or truncates the file at the specified path for writing.
func (f FileStore) Write() (io.WriteCloser, error) {
	return os.Create(string(f))
}
