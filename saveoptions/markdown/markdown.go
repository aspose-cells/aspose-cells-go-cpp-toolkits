// Package markdown provides a SaveOption and configuration options for
// exporting spreadsheets as Markdown tables.
package markdown

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	saveoptions "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// Config holds the native-typed option values for Markdown save options.
//
// Pointer fields carry presence semantics: a nil pointer means the option
// was never set (so the native default is kept), while a non-nil pointer
// means the caller explicitly requested the value, including zero/false.
type Config struct {
	encoding                   *asposecells.EncodingType
	formatStrategy             *asposecells.CellValueFormatStrategy
	lineSeparator              *string
	tableHeaderType            *asposecells.MarkdownTableHeaderType
	sheetSet                   *asposecells.SheetSet
	exportImagesAsBase64       *bool
	calculateFormula           *bool
	exportHyperlinkAsReference *bool
	alignColumnPadding         *byte
	splitTablesByBlankRow      *bool
	officeMathOutputType       *asposecells.HtmlOfficeMathOutputType
	clearData                  *bool
	cachedFileFolder           *string
	validateMergedAreas        *bool
	mergeAreas                 *bool
	createDirectory            *bool
	sortNames                  *bool
	sortExternalNames          *bool
	refreshChartCache          *bool
	checkExcelRestriction      *bool
	updateSmartArt             *bool
	encryptDocumentProperties  *bool
}

// Apply processes the given source byte slice as a Markdown file and returns the converted output.
// This method satisfies the saveoptions.SaveOption (or equivalent) interface, enabling Markdown-specific export logic.
//
// Parameters:
// - source: A byte slice representing the input spreadsheet or data source. The implementation may interpret
// this as an intermediate format (e.g., XLSX or CSV bytes) and convert it into Markdown format.
//
// Returns:
// - []byte: The resulting Markdown file content as a byte slice.
// - error: error information.
func (c *Config) Apply(source []byte) ([]byte, error) {
	opts, err := asposecells.NewMarkdownSaveOptions()
	if err != nil {
		return nil, err
	}
	if c.encoding != nil {
		if err := opts.SetEncoding(*c.encoding); err != nil {
			return nil, err
		}
	}
	if c.formatStrategy != nil {
		if err := opts.SetFormatStrategy(*c.formatStrategy); err != nil {
			return nil, err
		}
	}
	if c.lineSeparator != nil {
		if err := opts.SetLineSeparator(*c.lineSeparator); err != nil {
			return nil, err
		}
	}
	if c.tableHeaderType != nil {
		if err := opts.SetTableHeaderType(*c.tableHeaderType); err != nil {
			return nil, err
		}
	}
	if c.sheetSet != nil {
		if err := opts.SetSheetSet(c.sheetSet); err != nil {
			return nil, err
		}
	}
	if c.exportImagesAsBase64 != nil {
		if err := opts.SetExportImagesAsBase64(*c.exportImagesAsBase64); err != nil {
			return nil, err
		}
	}
	if c.calculateFormula != nil {
		if err := opts.SetCalculateFormula(*c.calculateFormula); err != nil {
			return nil, err
		}
	}
	if c.exportHyperlinkAsReference != nil {
		if err := opts.SetExportHyperlinkAsReference(*c.exportHyperlinkAsReference); err != nil {
			return nil, err
		}
	}
	if c.alignColumnPadding != nil {
		if err := opts.SetAlignColumnPadding(*c.alignColumnPadding); err != nil {
			return nil, err
		}
	}
	if c.splitTablesByBlankRow != nil {
		if err := opts.SetSplitTablesByBlankRow(*c.splitTablesByBlankRow); err != nil {
			return nil, err
		}
	}
	if c.officeMathOutputType != nil {
		if err := opts.SetOfficeMathOutputType(*c.officeMathOutputType); err != nil {
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
	return "md"
}

type Option func(*Config)

func init() {
	formats.Register("md", func() saveoptions.SaveOption {
		return New()
	})
}

// New creates a new instance of md save options
//
// The New function creates an instance of md SaveOption using the Functional Options Pattern. This function accepts a variable number of Option function parameters, and each Option function modifies the configuration of SaveOption.
//
// Parameters:
//
//	opts ... Option - A variable number of option functions used to configure SaveOption
//
// Return value:
// md SaveOption - Configured instance of the saved option
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

func WithEncoding(value asposecells.EncodingType) Option {
	return func(c *Config) {
		c.encoding = &value
	}
}
func WithFormatStrategy(value asposecells.CellValueFormatStrategy) Option {
	return func(c *Config) {
		c.formatStrategy = &value
	}
}
func WithLineSeparator(value string) Option {
	return func(c *Config) {
		c.lineSeparator = &value
	}
}

func WithTableHeaderType(value asposecells.MarkdownTableHeaderType) Option {
	return func(c *Config) {
		c.tableHeaderType = &value
	}
}
func WithSheetSet(value *asposecells.SheetSet) Option {
	return func(c *Config) {
		c.sheetSet = value
	}
}
func WithExportImagesAsBase64(value bool) Option {
	return func(c *Config) {
		c.exportImagesAsBase64 = &value
	}
}

func WithCalculateFormula(value bool) Option {
	return func(c *Config) {
		c.calculateFormula = &value
	}
}

func WithExportHyperlinkAsReference(value bool) Option {
	return func(c *Config) {
		c.exportHyperlinkAsReference = &value
	}
}

func WithAlignColumnPadding(value byte) Option {
	return func(c *Config) {
		c.alignColumnPadding = &value
	}
}

func WithSplitTablesByBlankRow(value bool) Option {
	return func(c *Config) {
		c.splitTablesByBlankRow = &value
	}
}

func WithOfficeMathOutputType(value asposecells.HtmlOfficeMathOutputType) Option {
	return func(c *Config) {
		c.officeMathOutputType = &value
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
