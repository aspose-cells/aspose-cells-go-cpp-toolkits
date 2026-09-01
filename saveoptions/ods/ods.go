// Package ods provides a SaveOption and configuration options for exporting
// spreadsheets as ODS (OpenDocument Spreadsheet) files.
package ods

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	saveoptions "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// Config holds the native-typed option values for Ods save options.
//
// Pointer fields carry presence semantics: a nil pointer means the option
// was never set (so the native default is kept), while a non-nil pointer
// means the caller explicitly requested the value, including zero/false.
type Config struct {
	generatorType     *asposecells.OdsGeneratorType
	odfStrictVersion  *asposecells.OpenDocumentFormatVersionType
	ignorePivotTables *bool
	saveoptions.CommonConfig
}

// Apply processes the given source byte slice as a Ods file and returns the converted output.
// This method satisfies the saveoptions.SaveOption (or equivalent) interface, enabling Ods-specific export logic.
//
// Parameters:
// - source: A byte slice representing the input spreadsheet or data source. The implementation may interpret
// this as an intermediate format (e.g., XLSX or CSV bytes) and convert it into Ods format.
//
// Returns:
// - []byte: The resulting Ods file content as a byte slice.
// - error: error information.

func (c *Config) Apply(source []byte) ([]byte, error) {
	opts, err := asposecells.NewOdsSaveOptions()
	if err != nil {
		return nil, err
	}

	if c.generatorType != nil {
		if err := opts.SetGeneratorType(*c.generatorType); err != nil {
			return nil, err
		}
	}
	if c.odfStrictVersion != nil {
		if err := opts.SetOdfStrictVersion(*c.odfStrictVersion); err != nil {
			return nil, err
		}
	}
	if c.ignorePivotTables != nil {
		if err := opts.SetIgnorePivotTables(*c.ignorePivotTables); err != nil {
			return nil, err
		}
	}
	if err := c.ApplyCommon(opts); err != nil {
		return nil, err
	}
	workbook, err := asposecells.NewWorkbook_Stream(source)
	if err != nil {
		return nil, err
	}
	saveOption := opts.ToSaveOptions()
	result, err := workbook.Save_SaveOptions(saveOption)
	if err != nil {
		return nil, err
	}
	return result, nil
}
func (c *Config) GetFormat() string {
	return "ods"
}

type Option func(*Config)

func init() {
	formats.Register("ods", func() saveoptions.SaveOption {
		return New()
	})
}

// New creates a new instance of ods save options
//
// The New function creates an instance of ods SaveOption using the Functional Options Pattern. This function accepts a variable number of Option function parameters, and each Option function modifies the configuration of SaveOption.
//
// Parameters:
//
//	opts ... Option - A variable number of option functions used to configure SaveOption
//
// Return value:
// ods SaveOption - Configured instance of the saved option
//
// Usage example:
//
// create default options
//
//	opts := New()
//
// create an instance with custom options
//
//	opts := New(
//	    WithCachedFileFolder("D:\\cached_folder"),
//	    WithClearData(true),
//
// )
//
// // use the option to perform the save operation
//
//	err := SaveFile(data, opts)
//
// Precautions:
// - If no options are provided, return the default configured SaveOption
// - Options are applied in the order provided, and the later applied options will overwrite the previous Settings All Option functions are thread-safe, but the SaveOption instance itself is not
//
// Related types:
//
//	type Option func(*Config)
//	type Config struct { ...  }
func New(opts ...Option) saveoptions.SaveOption {

	cfg := &Config{}

	for _, o := range opts {
		o(cfg)
	}

	return cfg
}

func WithGeneratorType(value asposecells.OdsGeneratorType) Option {
	return func(c *Config) {
		c.generatorType = &value
	}
}
func WithOdfStrictVersion(value asposecells.OpenDocumentFormatVersionType) Option {
	return func(c *Config) {
		c.odfStrictVersion = &value
	}
}
func WithIgnorePivotTables(value bool) Option {
	return func(c *Config) {
		c.ignorePivotTables = &value
	}
}

func WithClearData(value bool) Option {
	return func(c *Config) {
		c.ClearData = &value
	}
}

func WithCachedFileFolder(value string) Option {
	return func(c *Config) {
		c.CachedFileFolder = &value
	}
}

func WithValidateMergedAreas(value bool) Option {
	return func(c *Config) {
		c.ValidateMergedAreas = &value
	}
}

func WithMergeAreas(value bool) Option {
	return func(c *Config) {
		c.MergeAreas = &value
	}
}

func WithCreateDirectory(value bool) Option {
	return func(c *Config) {
		c.CreateDirectory = &value
	}
}

func WithSortNames(value bool) Option {
	return func(c *Config) {
		c.SortNames = &value
	}
}

func WithSortExternalNames(value bool) Option {
	return func(c *Config) {
		c.SortExternalNames = &value
	}
}

func WithRefreshChartCache(value bool) Option {
	return func(c *Config) {
		c.RefreshChartCache = &value
	}
}

func WithCheckExcelRestriction(value bool) Option {
	return func(c *Config) {
		c.CheckExcelRestriction = &value
	}
}

func WithUpdateSmartArt(value bool) Option {
	return func(c *Config) {
		c.UpdateSmartArt = &value
	}
}

func WithEncryptDocumentProperties(value bool) Option {
	return func(c *Config) {
		c.EncryptDocumentProperties = &value
	}
}
