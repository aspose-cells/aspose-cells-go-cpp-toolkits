package json

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	saveoptions "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// Config holds the native-typed option values for Json save options.
//
// Pointer fields carry presence semantics: a nil pointer means the option
// was never set (so the native default is kept), while a non-nil pointer
// means the caller explicitly requested the value, including zero/false.
type Config struct {
	exportStylePool           *bool
	exportHyperlinkType       *asposecells.JsonExportHyperlinkType
	skipEmptyRows             *bool
	sheetIndexes              []int32
	schemas                   []string
	exportArea                *asposecells.CellArea
	hasHeaderRow              *bool
	exportAsString            *bool
	indent                    *string
	exportNestedStructure     *bool
	exportEmptyCells          *bool
	alwaysExportAsJsonObject  *bool
	toExcelStruct             *bool
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

// Apply processes the given source byte slice as a Json file and returns the converted output.
// This method satisfies the saveoptions.SaveOption (or equivalent) interface, enabling Json-specific export logic.
//
// Parameters:
// - source: A byte slice representing the input spreadsheet or data source. The implementation may interpret
// this as an intermediate format (e.g., XLSX or CSV bytes) and convert it into Json format.
//
// Returns:
// - []byte: The resulting Json file content as a byte slice.
// - error: error information.
func (c *Config) Apply(source []byte) ([]byte, error) {
	opts, err := asposecells.NewJsonSaveOptions()
	if err != nil {
		return nil, err
	}
	if c.exportStylePool != nil {
		if err := opts.SetExportStylePool(*c.exportStylePool); err != nil {
			return nil, err
		}
	}
	if c.exportHyperlinkType != nil {
		if err := opts.SetExportHyperlinkType(*c.exportHyperlinkType); err != nil {
			return nil, err
		}
	}
	if c.skipEmptyRows != nil {
		if err := opts.SetSkipEmptyRows(*c.skipEmptyRows); err != nil {
			return nil, err
		}
	}
	if c.sheetIndexes != nil {
		if err := opts.SetSheetIndexes(c.sheetIndexes); err != nil {
			return nil, err
		}
	}
	if c.schemas != nil {
		if err := opts.SetSchemas(c.schemas); err != nil {
			return nil, err
		}
	}
	if c.exportArea != nil {
		if err := opts.SetExportArea(c.exportArea); err != nil {
			return nil, err
		}
	}
	if c.hasHeaderRow != nil {
		if err := opts.SetHasHeaderRow(*c.hasHeaderRow); err != nil {
			return nil, err
		}
	}
	if c.exportAsString != nil {
		if err := opts.SetExportAsString(*c.exportAsString); err != nil {
			return nil, err
		}
	}
	if c.indent != nil {
		if err := opts.SetIndent(*c.indent); err != nil {
			return nil, err
		}
	}
	if c.exportNestedStructure != nil {
		if err := opts.SetExportNestedStructure(*c.exportNestedStructure); err != nil {
			return nil, err
		}
	}
	if c.exportEmptyCells != nil {
		if err := opts.SetExportEmptyCells(*c.exportEmptyCells); err != nil {
			return nil, err
		}
	}
	if c.alwaysExportAsJsonObject != nil {
		if err := opts.SetAlwaysExportAsJsonObject(*c.alwaysExportAsJsonObject); err != nil {
			return nil, err
		}
	}
	if c.toExcelStruct != nil {
		if err := opts.SetToExcelStruct(*c.toExcelStruct); err != nil {
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
	return "json"
}

type Option func(*Config)

func init() {
	formats.Register("json", func() saveoptions.SaveOption {
		return New()
	})
}

// New creates a new instance of json save options
//
// The New function creates an instance of json SaveOption using the Functional Options Pattern. This function accepts a variable number of Option function parameters, and each Option function modifies the configuration of SaveOption.
//
// Parameters:
//
//	opts ... Option - A variable number of option functions used to configure SaveOption
//
// Return value:
// json SaveOption - Configured instance of the saved option
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

func WithExportStylePool(value bool) Option {
	return func(c *Config) {
		c.exportStylePool = &value
	}
}

func WithExportHyperlinkType(value asposecells.JsonExportHyperlinkType) Option {
	return func(c *Config) {
		c.exportHyperlinkType = &value
	}
}
func WithSkipEmptyRows(value bool) Option {
	return func(c *Config) {
		c.skipEmptyRows = &value
	}
}

func WithSheetIndexes(value []int32) Option {
	return func(c *Config) {
		c.sheetIndexes = value
	}
}
func WithSchemas(value []string) Option {
	return func(c *Config) {
		c.schemas = value
	}
}
func WithExportArea(value *asposecells.CellArea) Option {
	return func(c *Config) {
		c.exportArea = value
	}
}
func WithHasHeaderRow(value bool) Option {
	return func(c *Config) {
		c.hasHeaderRow = &value
	}
}

func WithExportAsString(value bool) Option {
	return func(c *Config) {
		c.exportAsString = &value
	}
}

func WithIndent(value string) Option {
	return func(c *Config) {
		c.indent = &value
	}
}

func WithExportNestedStructure(value bool) Option {
	return func(c *Config) {
		c.exportNestedStructure = &value
	}
}

func WithExportEmptyCells(value bool) Option {
	return func(c *Config) {
		c.exportEmptyCells = &value
	}
}

func WithAlwaysExportAsJsonObject(value bool) Option {
	return func(c *Config) {
		c.alwaysExportAsJsonObject = &value
	}
}

func WithToExcelStruct(value bool) Option {
	return func(c *Config) {
		c.toExcelStruct = &value
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
