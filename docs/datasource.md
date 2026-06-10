# datasource

## Types

### BytesSource

```go
type BytesSource []byte
```

BytesSource is an implementation of a data source based on in-memory byte slices. This is very useful in unit tests or when processing data that has been loaded into memory.

#### Methods

##### Open

```go
func (b BytesSource) Open 
```

Open wraps the byte slice into an io.ReadCloser. Since the data is in memory, this operation usually does not fail (returns nil error). io.NopCloser is used to wrap an io.Reader into an io.ReadCloser, and its Close method is an empty operation.

### DataSource

```go
type DataSource interface {
	Open() (io.ReadCloser, error)
}
```

The DataSource defines the general interface for data sources. Any type that implements this interface can serve as a data provider. The Open method is responsible for returning an io.ReadCloser, and the caller must close it after

### FilePathSource

```go
type FilePathSource string
```

FilePathSource is an implementation of a data source based on local file system paths. It implements the DataSource interface and is used to read files from the specified path.

#### Methods

##### Open

```go
func (p FilePathSource) Open 
```

Open Opens the file at the specified path. It directly calls os.Open, so if the file does not exist or does not have the required permissions, it will return the corresponding system error.

### ReaderSource

```go
type ReaderSource struct {
	// contains filtered or unexported fields
}
```

ReaderSource is an adapter for the existing io.ReadCloser. When you already have an open stream (such as an HTTP response body), but the function signature requires a DataSource, use this.

#### Methods

##### Open

```go
func (r ReaderSource) Open 
```

Open directly returns the reader held internally. Note: ReaderSource is usually used to wrap already opened streams, so the Open method itself does not return an error.

