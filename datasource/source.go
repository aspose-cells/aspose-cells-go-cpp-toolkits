package datasource

import (
	"bytes"
	"io"
	"os"
)

// The DataSource defines the general interface for data sources.
// Any type that implements this interface can serve as a data provider.
// The Open method is responsible for returning an io.ReadCloser, and the caller must close it after
type DataSource interface {
	Open() (io.ReadCloser, error)
}

// FilePathSource is an implementation of a data source based on local file system paths.
// It implements the DataSource interface and is used to read files from the specified path.
type FilePathSource string

// Open Opens the file at the specified path.
// It directly calls os.Open, so if the file does not exist or does not have the required permissions, it will return the corresponding system error.
func (p FilePathSource) Open() (io.ReadCloser, error) {
	return os.Open(string(p))
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

// ReaderSource is an adapter for the existing io.ReadCloser.
// When you already have an open stream (such as an HTTP response body), but the function signature requires a DataSource, use this.
type ReaderSource struct {
	reader io.ReadCloser
}

// Open directly returns the reader held internally.
// Note: ReaderSource is usually used to wrap already opened streams, so the Open method itself does not return an error.
func (r ReaderSource) Open() (io.ReadCloser, error) {
	return r.reader, nil
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
