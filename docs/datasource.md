# datasource

The `datasource` package abstracts spreadsheet input and output so toolkit entry points take `DataSource` / `DataSink` instead of concrete file, bytes, or stream types. One signature therefore covers every input and output shape, keeping the public API surface small.

## Input

### DataSource

```go
type DataSource interface {
	Open() (io.ReadCloser, error)
}
```

The `DataSource` interface defines a readable input. Any type that implements `Open` can serve as a data provider. `Open` returns an `io.ReadCloser` that the toolkit (via `internal/aspose/cells.ReadSource`) drains and closes.

### BytesSource

```go
type BytesSource []byte
```

`BytesSource` is a `DataSource` backed by an in-memory byte slice. Useful in unit tests or when processing data already loaded into memory. `Open` wraps the slice in an `io.ReadCloser` and never fails.

### FilePathSource

```go
type FilePathSource string
```

`FilePathSource` is a `DataSource` backed by a local file path. `Open` calls `os.Open`, so a missing file or a permissions problem surfaces the corresponding system error.

### ReaderSource

```go
func NewReaderSource(r io.Reader) *ReaderSource
```

`NewReaderSource` adapts an already-open stream (an HTTP response body, an open file, etc.) into a `DataSource`. The stream is consumed lazily on first use and buffered, so `Open` may be called repeatedly; every call returns a fresh reader over the same bytes. The caller remains responsible for closing the underlying stream.

## Output

### DataSink

```go
type DataSink interface {
	Write(name string, data []byte) error
}
```

The `DataSink` interface defines a writable output. The `name` parameter is used by multi-output operations (for example `manipulator.Split` writing one file or archive entry per worksheet); single-output sinks ignore it.

### FilePathSink

```go
type FilePathSink string
```

`FilePathSink` writes each `Write` call to the file at the given path, truncating any existing file. `name` is ignored.

### WriterSink

```go
func NewWriterSink(w io.Writer) *WriterSink
```

`NewWriterSink` wraps an `io.Writer` (a file, `bytes.Buffer`, HTTP response writer, …) as a `DataSink`. `name` is ignored.

### BytesSink

```go
type BytesSink struct {
	// contains unexported fields
}

func (b *BytesSink) Write(name string, data []byte) error
func (b *BytesSink) Bytes() []byte
```

`BytesSink` accumulates every `Write` call into an in-memory buffer. Hand a `*BytesSink` to any toolkit entry point when you want the result as bytes, then read it back with `Bytes()`.

### FolderSink

```go
type FolderSink string
```

`FolderSink` writes each `Write` call to a file named `name` inside the given folder, creating the folder if needed. It backs `manipulator.Split`'s per-worksheet file output.

### ZipSink

```go
func NewZipSink(zw *zip.Writer) *ZipSink
```

`NewZipSink` wraps a `zip.Writer` as a `DataSink`. Each `Write` call creates an entry named `name` in the archive. The caller closes the `zip.Writer` after the operation to finalize the archive. It backs `manipulator.Split`'s per-worksheet archive output.
