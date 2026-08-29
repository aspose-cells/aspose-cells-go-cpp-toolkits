package csv

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	saveoptions "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

type Config struct {
	separator                    *byte
	separatorString              *string
	encoding                     *asposecells.EncodingType
	quoteType                    *asposecells.TxtValueQuoteType
	formatStrategy               *asposecells.CellValueFormatStrategy
	trimLeadingBlankRowAndColumn *bool
	trimTailingBlankCells        *bool
	keepSeparatorsForBlankRow    *bool
	exportArea                   *asposecells.CellArea
	exportQuotePrefix            *bool
	exportAllSheets              *bool
	clearData                    *bool
	cachedFileFolder             *string
	validateMergedAreas          *bool
	mergeAreas                   *bool
	createDirectory              *bool
	sortNames                    *bool
	sortExternalNames            *bool
	refreshChartCache            *bool
	checkExcelRestriction        *bool
	updateSmartArt               *bool
	encryptDocumentProperties    *bool
}

// Apply processes the given source byte slice as a Csv file and returns the converted output.
// This method satisfies the saveoptions.SaveOption interface, enabling Csv-specific export logic.
//
// Parameters:
// - source: A byte slice representing the input spreadsheet or data source.
//
// Returns:
// - []byte: The resulting Csv file content as a byte slice.
// - error: error information.
func (c *Config) Apply(source []byte) ([]byte, error) {
	opts, err := asposecells.NewTxtSaveOptions_SaveFormat(asposecells.SaveFormat_Csv)
	if err != nil {
		return nil, err
	}
	if c.separator != nil {
		if err := opts.SetSeparator(*c.separator); err != nil {
			return nil, err
		}
	}
	if c.separatorString != nil {
		if err := opts.SetSeparatorString(*c.separatorString); err != nil {
			return nil, err
		}
	}
	if c.encoding != nil {
		if err := opts.SetEncoding(*c.encoding); err != nil {
			return nil, err
		}
	}
	if c.quoteType != nil {
		if err := opts.SetQuoteType(*c.quoteType); err != nil {
			return nil, err
		}
	}
	if c.formatStrategy != nil {
		if err := opts.SetFormatStrategy(*c.formatStrategy); err != nil {
			return nil, err
		}
	}
	if c.trimLeadingBlankRowAndColumn != nil {
		if err := opts.SetTrimLeadingBlankRowAndColumn(*c.trimLeadingBlankRowAndColumn); err != nil {
			return nil, err
		}
	}
	if c.trimTailingBlankCells != nil {
		if err := opts.SetTrimTailingBlankCells(*c.trimTailingBlankCells); err != nil {
			return nil, err
		}
	}
	if c.keepSeparatorsForBlankRow != nil {
		if err := opts.SetKeepSeparatorsForBlankRow(*c.keepSeparatorsForBlankRow); err != nil {
			return nil, err
		}
	}
	if c.exportArea != nil {
		if err := opts.SetExportArea(c.exportArea); err != nil {
			return nil, err
		}
	}
	if c.exportQuotePrefix != nil {
		if err := opts.SetExportQuotePrefix(*c.exportQuotePrefix); err != nil {
			return nil, err
		}
	}
	if c.exportAllSheets != nil {
		if err := opts.SetExportAllSheets(*c.exportAllSheets); err != nil {
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
	return workbook.Save_SaveOptions(saveOption)
}
func (c *Config) GetFormat() string {
	return "csv"
}

type Option func(*Config)

func init() {
	formats.Register("csv", func() saveoptions.SaveOption {
		return New()
	})
}

// New creates a new instance of csv save options using the Functional Options Pattern.
//
// Parameters:
//
//	opts ... Option - A variable number of option functions used to configure SaveOption
//
// Return value:
// csv SaveOption - Configured instance of the saved option
func New(opts ...Option) saveoptions.SaveOption {
	cfg := &Config{}
	for _, o := range opts {
		o(cfg)
	}
	return cfg
}

func WithSeparator(value byte) Option {
	return func(c *Config) {
		c.separator = &value
	}
}

func WithSeparatorString(value string) Option {
	return func(c *Config) {
		c.separatorString = &value
	}
}

func WithEncoding(value asposecells.EncodingType) Option {
	return func(c *Config) {
		c.encoding = &value
	}
}
func WithQuoteType(value asposecells.TxtValueQuoteType) Option {
	return func(c *Config) {
		c.quoteType = &value
	}
}
func WithFormatStrategy(value asposecells.CellValueFormatStrategy) Option {
	return func(c *Config) {
		c.formatStrategy = &value
	}
}
func WithTrimLeadingBlankRowAndColumn(value bool) Option {
	return func(c *Config) {
		c.trimLeadingBlankRowAndColumn = &value
	}
}

func WithTrimTailingBlankCells(value bool) Option {
	return func(c *Config) {
		c.trimTailingBlankCells = &value
	}
}

func WithKeepSeparatorsForBlankRow(value bool) Option {
	return func(c *Config) {
		c.keepSeparatorsForBlankRow = &value
	}
}

func WithExportArea(value *asposecells.CellArea) Option {
	return func(c *Config) {
		c.exportArea = value
	}
}
func WithExportQuotePrefix(value bool) Option {
	return func(c *Config) {
		c.exportQuotePrefix = &value
	}
}

func WithExportAllSheets(value bool) Option {
	return func(c *Config) {
		c.exportAllSheets = &value
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
