package datasource

import (
	"bytes"
	"io"
	"os"
)

type DataSource interface {
	Open() (io.ReadCloser, error)
}

type FilePathSource string

func (p FilePathSource) Open() (io.ReadCloser, error) {
	return os.Open(string(p))
}

type BytesSource []byte

func (b BytesSource) Open() (io.ReadCloser, error) {
	return io.NopCloser(bytes.NewReader(b)), nil
}

type ReaderSource struct {
	reader io.ReadCloser
}

func (r ReaderSource) Open() (io.ReadCloser, error) {
	return r.reader, nil
}
