// Package formats maps file extensions to saveoptions.SaveOption factories.
//
// Implementations register themselves by extension so callers can resolve an
// output format from a filename. The registry is concurrency-safe: Register,
// Unregister, Get, and List may be called from multiple goroutines, and List
// returns a sorted, stable snapshot.
package formats

import (
	"sort"
	"strings"
	"sync"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
)

var (
	registryMu sync.RWMutex
	registry   = make(map[string]func() saveoptions.SaveOption)
)

// Register maps a file extension to a factory that produces its SaveOption.
// The extension is normalized to lower case and lookups are case-insensitive.
// Register is safe for concurrent use and may be called at any time, not only
// during package initialization.
func Register(ext string, factory func() saveoptions.SaveOption) {
	registryMu.Lock()
	defer registryMu.Unlock()
	registry[strings.ToLower(ext)] = factory
}

// Unregister removes a previously registered extension. It is safe for
// concurrent use and is a no-op if the extension is not registered.
func Unregister(ext string) {
	registryMu.Lock()
	defer registryMu.Unlock()
	delete(registry, strings.ToLower(ext))
}

// Get returns a new SaveOption for the given extension, or nil if the
// extension is not registered. Get is safe for concurrent use.
func Get(ext string) saveoptions.SaveOption {
	registryMu.RLock()
	factory, ok := registry[strings.ToLower(ext)]
	registryMu.RUnlock()
	if !ok {
		return nil
	}
	return factory()
}

// List returns all registered extensions in sorted order. List is safe for
// concurrent use.
func List() []string {
	registryMu.RLock()
	defer registryMu.RUnlock()
	exts := make([]string, 0, len(registry))
	for ext := range registry {
		exts = append(exts, ext)
	}
	sort.Strings(exts)
	return exts
}
