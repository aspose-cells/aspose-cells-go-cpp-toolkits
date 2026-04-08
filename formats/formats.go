package formats

import (
	"strings"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions"
)

var registry = make(map[string]func() saveoptions.SaveOption)

func Register(ext string, factory func() saveoptions.SaveOption) {
	registry[strings.ToLower(ext)] = factory
}

func Get(ext string) saveoptions.SaveOption {
	if factory, ok := registry[strings.ToLower(ext)]; ok {
		return factory()
	}
	return nil
}

func List() []string {
	exts := make([]string, 0, len(registry))
	for ext := range registry {
		exts = append(exts, ext)
	}
	return exts
}
