// Package formats maps file extensions and Aspose.Cells format types to
// saveoptions.SaveOption factories.
//
// Implementations register themselves by extension so callers can resolve an
// output format from a filename or from a FileFormatType. The registry is
// concurrency-safe: Register, Unregister, Get, and List may be called from
// multiple goroutines, and List returns a sorted, stable snapshot.
package formats

import (
	"fmt"
	"sort"
	"strings"
	"sync"

	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
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

// FileFormatToSaveFormat maps an engine input format to the SaveFormat that
// writes the same format back.
//
// A format with no matching output — one the toolkit does not support saving,
// or an unrecognized value — is reported as ErrUnsupportedFormat rather than
// mapped to SaveFormat_Auto. Auto leaves the choice to the engine, which does
// not produce the format the caller asked for: a round-trip save of a document
// loaded in such a format would silently come back as a different format.
func FileFormatToSaveFormat(formatType asposecells.FileFormatType) (asposecells.SaveFormat, error) {
	switch formatType {
	case asposecells.FileFormatType_Azw3:
		return asposecells.SaveFormat_Azw3, nil
	case asposecells.FileFormatType_Bmp:
		return asposecells.SaveFormat_Bmp, nil
	case asposecells.FileFormatType_Csv:
		return asposecells.SaveFormat_Csv, nil
	case asposecells.FileFormatType_Docx,
		asposecells.FileFormatType_Docm,
		asposecells.FileFormatType_Doc,
		asposecells.FileFormatType_Dotm,
		asposecells.FileFormatType_Rtf:
		return asposecells.SaveFormat_Docx, nil
	case asposecells.FileFormatType_Dif:
		return asposecells.SaveFormat_Dif, nil
	case asposecells.FileFormatType_Dbf:
		return asposecells.SaveFormat_Dbf, nil
	case asposecells.FileFormatType_Epub:
		return asposecells.SaveFormat_Epub, nil
	case asposecells.FileFormatType_Emf:
		return asposecells.SaveFormat_Emf, nil
	case asposecells.FileFormatType_Excel97To2003:
		return asposecells.SaveFormat_Excel97To2003, nil
	case asposecells.FileFormatType_Fods:
		return asposecells.SaveFormat_Fods, nil
	case asposecells.FileFormatType_Gif:
		return asposecells.SaveFormat_Gif, nil
	case asposecells.FileFormatType_Html:
		return asposecells.SaveFormat_Html, nil
	case asposecells.FileFormatType_MHtml:
		return asposecells.SaveFormat_MHtml, nil
	case asposecells.FileFormatType_Json:
		return asposecells.SaveFormat_Json, nil
	case asposecells.FileFormatType_Jpg:
		return asposecells.SaveFormat_Jpg, nil
	case asposecells.FileFormatType_Markdown:
		return asposecells.SaveFormat_Markdown, nil
	case asposecells.FileFormatType_Numbers35:
		return asposecells.SaveFormat_Numbers, nil
	case asposecells.FileFormatType_Ods:
		return asposecells.SaveFormat_Ods, nil
	case asposecells.FileFormatType_Ots:
		return asposecells.SaveFormat_Ots, nil
	case asposecells.FileFormatType_Png:
		return asposecells.SaveFormat_Png, nil
	case asposecells.FileFormatType_Pdf:
		return asposecells.SaveFormat_Pdf, nil
	case asposecells.FileFormatType_Ppt,
		asposecells.FileFormatType_Pptx,
		asposecells.FileFormatType_Ppsm:
		return asposecells.SaveFormat_Pptx, nil
	case asposecells.FileFormatType_Sxc:
		return asposecells.SaveFormat_Sxc, nil
	case asposecells.FileFormatType_Svg:
		return asposecells.SaveFormat_Svg, nil
	case asposecells.FileFormatType_SqlScript:
		return asposecells.SaveFormat_SqlScript, nil
	case asposecells.FileFormatType_SpreadsheetML:
		return asposecells.SaveFormat_SpreadsheetML, nil
	case asposecells.FileFormatType_Tsv:
		return asposecells.SaveFormat_Tsv, nil
	case asposecells.FileFormatType_Tiff:
		return asposecells.SaveFormat_Tiff, nil
	case asposecells.FileFormatType_Xlsm:
		return asposecells.SaveFormat_Xlsm, nil
	case asposecells.FileFormatType_Xlsx:
		return asposecells.SaveFormat_Xlsx, nil
	case asposecells.FileFormatType_Xlsb:
		return asposecells.SaveFormat_Xlsb, nil
	case asposecells.FileFormatType_Xlam:
		return asposecells.SaveFormat_Xlam, nil
	case asposecells.FileFormatType_Xlt:
		return asposecells.SaveFormat_Xlt, nil
	case asposecells.FileFormatType_Xltx:
		return asposecells.SaveFormat_Xltx, nil
	case asposecells.FileFormatType_Xml:
		return asposecells.SaveFormat_Xml, nil
	case asposecells.FileFormatType_Xps:
		return asposecells.SaveFormat_Xps, nil
	}

	return asposecells.SaveFormat_Auto,
		fmt.Errorf("%w: file format type %d has no save format", toolkiterrors.ErrUnsupportedFormat, formatType)
}
