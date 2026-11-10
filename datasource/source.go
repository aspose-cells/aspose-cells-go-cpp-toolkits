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
	"fmt"
	"io"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"sync"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
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
	err    error
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
	// The stream was consumed by an earlier call that failed. Report the same
	// error again instead of silently serving an empty source: a caller may
	// re-read a source after a failed load, and a second attempt must not look
	// like a successful read of zero bytes.
	if r.err != nil {
		return nil, r.err
	}
	if r.reader == nil {
		return nil, nil
	}
	data, err := io.ReadAll(r.reader)
	r.reader = nil
	if err != nil {
		r.err = err
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

// MultiOutputSink is a DataSink that accepts one named output per Write call
// and keeps them all, rather than one output total. FolderSink and ZipSink
// implement it; FilePathSink, WriterSink, and BytesSink do not.
//
// manipulator.Split writes one file per worksheet, so it requires this
// interface: a single-output sink would accept every Write and silently keep
// only the last worksheet. A custom sink that writes one output per name
// implements MultiOutput with a no-op method to opt in.
type MultiOutputSink interface {
	DataSink
	// MultiOutput marks the sink as accepting multiple named outputs.
	MultiOutput()
}

// FilePathSink writes each Write call to a file at the given path. name is
// ignored.
type FilePathSink string

// Write writes data to the file at the sink's path, truncating any existing
// file. Missing parent directories are created first, matching FolderSink, so
// writing to a not-yet-existing output folder does not fail.
func (p FilePathSink) Write(_ string, data []byte) error {
	dir := filepath.Dir(string(p))
	if dir != "" && dir != "." {
		if err := os.MkdirAll(dir, 0o755); err != nil {
			return fmt.Errorf("create output directory %q: %w", dir, err)
		}
	}
	if err := writeFileAtomic(string(p), data); err != nil {
		return fmt.Errorf("write %q: %w", string(p), err)
	}
	return nil
}

// writeFileAtomic writes data to a temporary file in the destination directory
// and renames it into place. A crash or a full disk therefore leaves the
// previous file intact instead of a truncated one at a path a later run may
// trust.
func writeFileAtomic(path string, data []byte) error {
	dir := filepath.Dir(path)
	tmp, err := os.CreateTemp(dir, "."+filepath.Base(path)+".tmp*")
	if err != nil {
		return err
	}
	tmpName := tmp.Name()
	cleanup := func() { _ = os.Remove(tmpName) }

	if _, err := tmp.Write(data); err != nil {
		_ = tmp.Close()
		cleanup()
		return err
	}
	if err := tmp.Close(); err != nil {
		cleanup()
		return err
	}
	if err := os.Chmod(tmpName, 0o644); err != nil {
		cleanup()
		return err
	}
	if err := os.Rename(tmpName, path); err != nil {
		cleanup()
		return err
	}
	return nil
}

// WriterSink writes each Write call to the wrapped io.Writer. name is ignored.
type WriterSink struct {
	w io.Writer
}

// NewWriterSink wraps an io.Writer as a DataSink.
func NewWriterSink(w io.Writer) *WriterSink {
	return &WriterSink{w: w}
}

// Write forwards data to the wrapped writer. A short write with a nil error is
// reported as io.ErrShortWrite rather than treated as success.
func (s WriterSink) Write(_ string, data []byte) error {
	n, err := s.w.Write(data)
	if err != nil {
		return err
	}
	if n != len(data) {
		return io.ErrShortWrite
	}
	return nil
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

// SafeOutputName validates the name a multi-output sink was asked to write
// under, and is the check FolderSink and ZipSink apply before writing.
//
// That name is not the caller's own: Split, and anything else that writes one
// output per worksheet, passes the source document's worksheet name, so it is
// untrusted input. A name carrying a path separator, a volume prefix, or a ".."
// segment would let the document being read decide where the output lands —
// writing outside the sink's folder, or, for the archive sink, creating a
// zip-slip entry that escapes wherever the archive is later extracted.
//
// The file format already forbids these characters in a worksheet name, but the
// check is repeated here rather than taken on the engine's word: the engine's
// validation is not a security boundary the toolkit controls, and a miss is an
// arbitrary file write.
//
// Exported because a custom DataSink that writes one output per name needs the
// same check, and duplicating it invites the two copies to drift apart.
func SafeOutputName(name string) error {
	return safeOutputName(name, true)
}

// safeArchiveName applies the traversal checks to a zip entry name. Device
// names and trailing dots are a filesystem concern, not an archive one, so an
// entry named "CON" is legal even on Windows; an entry that escapes the
// extraction directory is not.
func safeArchiveName(name string) error {
	return safeOutputName(name, false)
}

func safeOutputName(name string, filesystem bool) error {
	if name == "" {
		return fmt.Errorf("%w: name is empty", toolkiterrors.ErrUnsafeSinkName)
	}
	// A separator in either flavor — the name may have been produced on another
	// platform — a drive-relative "c:" (also an NTFS alternate data stream), or
	// an absolute path.
	if name != filepath.Base(name) ||
		strings.ContainsAny(name, `/\:`) ||
		filepath.IsAbs(name) ||
		filepath.VolumeName(name) != "" {
		return fmt.Errorf("name %q is a path, not a plain file name: %w", name, toolkiterrors.ErrUnsafeSinkName)
	}
	if name == "." || name == ".." {
		return fmt.Errorf("name %q denotes a directory: %w", name, toolkiterrors.ErrUnsafeSinkName)
	}
	if filesystem && runtime.GOOS == "windows" {
		if err := checkWindowsFileName(name); err != nil {
			return err
		}
	}
	return nil
}

// checkWindowsFileName rejects names Windows resolves to something other than
// the requested file. The device names are reserved with or without an
// extension ("CON.xlsx" opens the console), and a trailing dot or space is
// silently stripped by the filesystem, so "A. " and "A" name the same file.
func checkWindowsFileName(name string) error {
	if strings.HasSuffix(name, ".") || strings.HasSuffix(name, " ") {
		return fmt.Errorf("name %q ends with a dot or space, which Windows strips: %w", name, toolkiterrors.ErrUnsafeSinkName)
	}
	stem := name
	if i := strings.IndexByte(stem, '.'); i >= 0 {
		stem = stem[:i]
	}
	stem = strings.ToUpper(stem)
	switch stem {
	case "CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$":
		return fmt.Errorf("name %q is a reserved Windows device name: %w", name, toolkiterrors.ErrUnsafeSinkName)
	}
	if len(stem) == 4 &&
		(strings.HasPrefix(stem, "COM") || strings.HasPrefix(stem, "LPT")) &&
		stem[3] >= '1' && stem[3] <= '9' {
		return fmt.Errorf("name %q is a reserved Windows device name: %w", name, toolkiterrors.ErrUnsafeSinkName)
	}
	return nil
}

// FolderSink writes each Write call to a file named name inside the given
// folder, creating the folder if needed. It backs manipulator.Split's
// per-worksheet file output.
type FolderSink string

// Write creates the folder if needed and writes data to
// <folder>/<name>. A name that is not a plain file name — one containing a path
// separator or a ".." segment — is rejected with ErrUnsafeSinkName rather than
// resolved against the folder, since the name comes from the source document.
// MultiOutput marks FolderSink as a MultiOutputSink: each Write creates another
// file in the folder.
func (f FolderSink) MultiOutput() {}

func (f FolderSink) Write(name string, data []byte) error {
	if err := SafeOutputName(name); err != nil {
		return err
	}
	if err := os.MkdirAll(string(f), 0o755); err != nil {
		return fmt.Errorf("create output folder %q: %w", string(f), err)
	}
	path := filepath.Join(string(f), name)
	if err := writeFileAtomic(path, data); err != nil {
		return fmt.Errorf("write %q: %w", path, err)
	}
	return nil
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

// Write creates an entry named name in the archive and writes data to it. A
// name that is not a plain file name — one containing a path separator or a
// ".." segment — is rejected with ErrUnsafeSinkName, so the archive cannot
// carry a zip-slip entry built from a worksheet name.
// MultiOutput marks ZipSink as a MultiOutputSink: each Write creates another
// archive entry.
func (z ZipSink) MultiOutput() {}

func (z ZipSink) Write(name string, data []byte) error {
	if err := safeArchiveName(name); err != nil {
		return err
	}
	entry, err := z.zw.Create(name)
	if err != nil {
		return err
	}
	n, err := entry.Write(data)
	if err != nil {
		return err
	}
	if n != len(data) {
		return io.ErrShortWrite
	}
	return nil
}
