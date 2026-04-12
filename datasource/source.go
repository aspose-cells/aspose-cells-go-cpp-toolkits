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
