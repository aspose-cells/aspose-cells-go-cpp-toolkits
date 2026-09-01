// Package ooxml provides a SaveOption and configuration options for exporting
// spreadsheets as Office Open XML workbooks (XLSX, XLSM, XLTX, XLTM).
package ooxml

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	saveoptions "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// Config holds the native-typed option values for Ooxml save options.
//
// Pointer fields carry presence semantics: a nil pointer means the option
// was never set (so the native default is kept), while a non-nil pointer
// means the caller explicitly requested the value, including zero/false.
// format discriminates the OOXML sub-format: "", "xlsx", "xlsm", "xltx", or
// "xltm". Empty means the default (XLSX).
type Config struct {
	format                string
	asFlatOpc             *bool
	exportCellName        *bool
	updateZoom            *bool
	enableZip64           *bool
	embedOoxmlAsOleObject *bool
	compressionType       *asposecells.OoxmlCompressionType
	saveoptions.CommonConfig
}

// saveFormat maps the configured sub-format to the engine's SaveFormat.
func (c *Config) saveFormat() asposecells.SaveFormat {
	switch c.format {
	case "xlsm":
		return asposecells.SaveFormat_Xlsm
	case "xltx":
		return asposecells.SaveFormat_Xltx
	case "xltm":
		return asposecells.SaveFormat_Xltm
	default:
		return asposecells.SaveFormat_Xlsx
	}
}

// Apply processes the given source byte slice as an Ooxml file and returns the converted output.
// This method satisfies the saveoptions.SaveOption (or equivalent) interface, enabling Ooxml-specific export logic.
//
// Parameters:
// - source: A byte slice representing the input spreadsheet or data source. The implementation may interpret
// this as an intermediate format (e.g., XLSX or CSV bytes) and convert it into Ooxml format(Xlsx, Xlsm, and so on).
//
// Returns:
// - []byte: The resulting Ooxml file content as a byte slice.
// - error: error information.
func (c *Config) Apply(source []byte) ([]byte, error) {
	var opts *asposecells.OoxmlSaveOptions
	var err error
	if c.format != "" {
		opts, err = asposecells.NewOoxmlSaveOptions_SaveFormat(c.saveFormat())
	} else {
		opts, err = asposecells.NewOoxmlSaveOptions()
	}
	if err != nil {
		return nil, err
	}

	if c.asFlatOpc != nil {
		if err := opts.SetAsFlatOpc(*c.asFlatOpc); err != nil {
			return nil, err
		}
	}
	if c.exportCellName != nil {
		if err := opts.SetExportCellName(*c.exportCellName); err != nil {
			return nil, err
		}
	}
	if c.updateZoom != nil {
		if err := opts.SetUpdateZoom(*c.updateZoom); err != nil {
			return nil, err
		}
	}
	if c.enableZip64 != nil {
		if err := opts.SetEnableZip64(*c.enableZip64); err != nil {
			return nil, err
		}
	}
	if c.embedOoxmlAsOleObject != nil {
		if err := opts.SetEmbedOoxmlAsOleObject(*c.embedOoxmlAsOleObject); err != nil {
			return nil, err
		}
	}
	if c.compressionType != nil {
		if err := opts.SetCompressionType(*c.compressionType); err != nil {
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
	if c.format == "" {
		return "xlsx"
	}
	return c.format
}

type Option func(*Config)

func init() {
	formats.Register("xlsx", func() saveoptions.SaveOption {
		return New(WithFormat("xlsx"))
	})
	formats.Register("xlsm", func() saveoptions.SaveOption {
		return New(WithFormat("xlsm"))
	})
	formats.Register("xltx", func() saveoptions.SaveOption {
		return New(WithFormat("xltx"))
	})
	formats.Register("xltm", func() saveoptions.SaveOption {
		return New(WithFormat("xltm"))
	})
}

// WithFormat selects the OOXML sub-format to emit: "xlsx" (default), "xlsm",
// "xltx", or "xltm". The formats registry uses it, so converter.Convert with
// formats.Get("xlsm") actually produces an XLSM file.
func WithFormat(value string) Option {
	return func(c *Config) {
		c.format = value
	}
}

// New creates a new instance of xlsx save options
//
// The New function creates an instance of xlsx SaveOption using the Functional Options Pattern. This function accepts a variable number of Option function parameters, and each Option function modifies the configuration of SaveOption.
//
// Parameters:
//
//	opts ... Option - A variable number of option functions used to configure SaveOption
//
// Return value:
// xlsx SaveOption - Configured instance of the saved option
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

func WithAsFlatOpc(value bool) Option {
	return func(c *Config) {
		c.asFlatOpc = &value
	}
}

func WithExportCellName(value bool) Option {
	return func(c *Config) {
		c.exportCellName = &value
	}
}

func WithUpdateZoom(value bool) Option {
	return func(c *Config) {
		c.updateZoom = &value
	}
}

func WithEnableZip64(value bool) Option {
	return func(c *Config) {
		c.enableZip64 = &value
	}
}

func WithEmbedOoxmlAsOleObject(value bool) Option {
	return func(c *Config) {
		c.embedOoxmlAsOleObject = &value
	}
}

func WithCompressionType(value asposecells.OoxmlCompressionType) Option {
	return func(c *Config) {
		c.compressionType = &value
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
