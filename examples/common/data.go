// Package examples provides helpers shared by the example commands. It lives
// under examples/ (not internal/) so example-only code stays out of the library
// packages shipped with the toolkit.
package examples

import (
	"path/filepath"
	"runtime"
)

// root resolves parts relative to the module root using the path of this source
// file, so the helpers stay correct no matter which directory an example
// command is launched from (module root, an example directory, or CI).
func root(parts ...string) string {
	_, thisFile, _, ok := runtime.Caller(0)
	if !ok {
		return filepath.Join(parts...)
	}
	return filepath.Join(append([]string{filepath.Dir(thisFile), "..", ".."}, parts...)...)
}

// DataPath returns the absolute path to a file in the examples/data directory,
// regardless of the current working directory. Examples can therefore be run
// with `go run ./examples/...` from the module root or from inside an example
// directory without breaking their relative data references.
func DataPath(name string) string {
	return root("examples", "data", name)
}

// OutDir returns the absolute path to the given example's output directory
// (examples/<name>/out). Writing to an absolute path keeps an example's
// artifacts inside its own directory no matter where the command was launched
// from, so a `go run` from the module root never pollutes it.
func OutDir(example string) string {
	return root("examples", example, "out")
}

// OutPath returns the absolute path to a single output file inside the given
// example's out/ directory.
func OutPath(example, name string) string {
	return root("examples", example, "out", name)
}
