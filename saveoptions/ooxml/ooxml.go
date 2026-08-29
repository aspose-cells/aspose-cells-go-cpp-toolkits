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
type Config struct {
	asFlatOpc                 *bool
	exportCellName            *bool
	updateZoom                *bool
	enableZip64               *bool
	embedOoxmlAsOleObject     *bool
	compressionType           *asposecells.OoxmlCompressionType
	clearData                 *bool
	cachedFileFolder          *string
	validateMergedAreas       *bool
	mergeAreas                *bool
	createDirectory           *bool
	sortNames                 *bool
	sortExternalNames         *bool
	refreshChartCache         *bool
	checkExcelRestriction     *bool
	updateSmartArt            *bool
	encryptDocumentProperties *bool
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
	opts, err := asposecells.NewOoxmlSaveOptions()
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
	if c.clearData != nil {
		if err := opts.SetClearData(*c.clearData); err != nil {
			return nil, err
		}
	}
	if c.cachedFileFolder != nil {
		if err := opts.SetCachedFileFolder(*c.cachedFileFolder); err != nil {
			return nil, err
		}
	}
	if c.validateMergedAreas != nil {
		if err := opts.SetValidateMergedAreas(*c.validateMergedAreas); err != nil {
			return nil, err
		}
	}
	if c.mergeAreas != nil {
		if err := opts.SetMergeAreas(*c.mergeAreas); err != nil {
			return nil, err
		}
	}
	if c.createDirectory != nil {
		if err := opts.SetCreateDirectory(*c.createDirectory); err != nil {
			return nil, err
		}
	}
	if c.sortNames != nil {
		if err := opts.SetSortNames(*c.sortNames); err != nil {
			return nil, err
		}
	}
	if c.sortExternalNames != nil {
		if err := opts.SetSortExternalNames(*c.sortExternalNames); err != nil {
			return nil, err
		}
	}
	if c.refreshChartCache != nil {
		if err := opts.SetRefreshChartCache(*c.refreshChartCache); err != nil {
			return nil, err
		}
	}
	if c.checkExcelRestriction != nil {
		if err := opts.SetCheckExcelRestriction(*c.checkExcelRestriction); err != nil {
			return nil, err
		}
	}
	if c.updateSmartArt != nil {
		if err := opts.SetUpdateSmartArt(*c.updateSmartArt); err != nil {
			return nil, err
		}
	}
	if c.encryptDocumentProperties != nil {
		if err := opts.SetEncryptDocumentProperties(*c.encryptDocumentProperties); err != nil {
			return nil, err
		}
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
	return "xlsx"
}

type Option func(*Config)

func init() {
	formats.Register("xlsx", func() saveoptions.SaveOption {
		return New()
	})
	formats.Register("xlsm", func() saveoptions.SaveOption {
		return New()
	})
	formats.Register("xltx", func() saveoptions.SaveOption {
		return New()
	})
	formats.Register("xltm", func() saveoptions.SaveOption {
		return New()
	})
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
//	    WithExportAsString(true),
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
		c.clearData = &value
	}
}

func WithCachedFileFolder(value string) Option {
	return func(c *Config) {
		c.cachedFileFolder = &value
	}
}

func WithValidateMergedAreas(value bool) Option {
	return func(c *Config) {
		c.validateMergedAreas = &value
	}
}

func WithMergeAreas(value bool) Option {
	return func(c *Config) {
		c.mergeAreas = &value
	}
}

func WithCreateDirectory(value bool) Option {
	return func(c *Config) {
		c.createDirectory = &value
	}
}

func WithSortNames(value bool) Option {
	return func(c *Config) {
		c.sortNames = &value
	}
}

func WithSortExternalNames(value bool) Option {
	return func(c *Config) {
		c.sortExternalNames = &value
	}
}

func WithRefreshChartCache(value bool) Option {
	return func(c *Config) {
		c.refreshChartCache = &value
	}
}

func WithCheckExcelRestriction(value bool) Option {
	return func(c *Config) {
		c.checkExcelRestriction = &value
	}
}

func WithUpdateSmartArt(value bool) Option {
	return func(c *Config) {
		c.updateSmartArt = &value
	}
}

func WithEncryptDocumentProperties(value bool) Option {
	return func(c *Config) {
		c.encryptDocumentProperties = &value
	}
}
