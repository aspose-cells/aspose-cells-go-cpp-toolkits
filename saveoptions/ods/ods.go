package ods

import (
	"strconv"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/formats"
	. "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

type Config struct {
	generatorType             string
	odfStrictVersion          string
	ignorePivotTables         string
	clearData                 string
	cachedFileFolder          string
	validateMergedAreas       string
	mergeAreas                string
	createDirectory           string
	sortNames                 string
	sortExternalNames         string
	refreshChartCache         string
	checkExcelRestriction     string
	updateSmartArt            string
	encryptDocumentProperties string
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
	opts, _ := asposecells.NewOdsSaveOptions()

	if v, err := strconv.ParseInt(c.generatorType, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToOdsGeneratorType(int32(v)); err2 == nil {
			opts.SetGeneratorType(vv)
		}
	}

	if v, err := strconv.ParseInt(c.odfStrictVersion, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToOpenDocumentFormatVersionType(int32(v)); err2 == nil {
			opts.SetOdfStrictVersion(vv)
		}
	}

	if len(c.ignorePivotTables) > 0 {
		if v, err := strconv.ParseBool(c.ignorePivotTables); err == nil {
			opts.SetIgnorePivotTables(v)
		}
	}
	if len(c.clearData) > 0 {
		if v, err := strconv.ParseBool(c.clearData); err == nil {
			opts.SetClearData(v)
		}
	}
	if len(c.cachedFileFolder) > 0 {
		opts.SetCachedFileFolder(c.cachedFileFolder)
	}
	if len(c.validateMergedAreas) > 0 {
		if v, err := strconv.ParseBool(c.validateMergedAreas); err == nil {
			opts.SetValidateMergedAreas(v)
		}
	}
	if len(c.mergeAreas) > 0 {
		if v, err := strconv.ParseBool(c.mergeAreas); err == nil {
			opts.SetMergeAreas(v)
		}
	}
	if len(c.createDirectory) > 0 {
		if v, err := strconv.ParseBool(c.createDirectory); err == nil {
			opts.SetCreateDirectory(v)
		}
	}
	if len(c.sortNames) > 0 {
		if v, err := strconv.ParseBool(c.sortNames); err == nil {
			opts.SetSortNames(v)
		}
	}
	if len(c.sortExternalNames) > 0 {
		if v, err := strconv.ParseBool(c.sortExternalNames); err == nil {
			opts.SetSortExternalNames(v)
		}
	}
	if len(c.refreshChartCache) > 0 {
		if v, err := strconv.ParseBool(c.refreshChartCache); err == nil {
			opts.SetRefreshChartCache(v)
		}
	}
	if len(c.checkExcelRestriction) > 0 {
		if v, err := strconv.ParseBool(c.checkExcelRestriction); err == nil {
			opts.SetCheckExcelRestriction(v)
		}
	}
	if len(c.updateSmartArt) > 0 {
		if v, err := strconv.ParseBool(c.updateSmartArt); err == nil {
			opts.SetUpdateSmartArt(v)
		}
	}
	if len(c.encryptDocumentProperties) > 0 {
		if v, err := strconv.ParseBool(c.encryptDocumentProperties); err == nil {
			opts.SetEncryptDocumentProperties(v)
		}
	}
	workbook, _ := asposecells.NewWorkbook_Stream(source)
	saveOption := opts.ToSaveOptions()
	result, _ := workbook.Save_SaveOptions(saveOption)
	return result, nil
}
func (c *Config) GetFormat() string {
	return "ods"
}

type Option func(*Config)

func init() {
	formats.Register("ods", func() SaveOption {
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
func New(opts ...Option) SaveOption {

	cfg := &Config{}

	for _, o := range opts {
		o(cfg)
	}

	return cfg
}

func WithGeneratorType(value asposecells.OdsGeneratorType) Option {
	return func(c *Config) {
		c.generatorType = strconv.FormatInt(int64(value), 10)
	}
}
func WithOdfStrictVersion(value asposecells.OpenDocumentFormatVersionType) Option {
	return func(c *Config) {
		c.odfStrictVersion = strconv.FormatInt(int64(value), 10)
	}
}
func WithIgnorePivotTables(value bool) Option {
	return func(c *Config) {
		c.ignorePivotTables = strconv.FormatBool(value)
	}
}

func WithClearData(value bool) Option {
	return func(c *Config) {
		c.clearData = strconv.FormatBool(value)
	}
}

func WithCachedFileFolder(value string) Option {
	return func(c *Config) {
		c.cachedFileFolder = value
	}
}

func WithValidateMergedAreas(value bool) Option {
	return func(c *Config) {
		c.validateMergedAreas = strconv.FormatBool(value)
	}
}

func WithMergeAreas(value bool) Option {
	return func(c *Config) {
		c.mergeAreas = strconv.FormatBool(value)
	}
}

func WithCreateDirectory(value bool) Option {
	return func(c *Config) {
		c.createDirectory = strconv.FormatBool(value)
	}
}

func WithSortNames(value bool) Option {
	return func(c *Config) {
		c.sortNames = strconv.FormatBool(value)
	}
}

func WithSortExternalNames(value bool) Option {
	return func(c *Config) {
		c.sortExternalNames = strconv.FormatBool(value)
	}
}

func WithRefreshChartCache(value bool) Option {
	return func(c *Config) {
		c.refreshChartCache = strconv.FormatBool(value)
	}
}

func WithCheckExcelRestriction(value bool) Option {
	return func(c *Config) {
		c.checkExcelRestriction = strconv.FormatBool(value)
	}
}

func WithUpdateSmartArt(value bool) Option {
	return func(c *Config) {
		c.updateSmartArt = strconv.FormatBool(value)
	}
}

func WithEncryptDocumentProperties(value bool) Option {
	return func(c *Config) {
		c.encryptDocumentProperties = strconv.FormatBool(value)
	}
}
