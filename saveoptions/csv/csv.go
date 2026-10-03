// Package csv provides a SaveOption and configuration options for exporting
// spreadsheets as CSV files.
package csv

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	engine "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/engine"
	enums "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/enums"
	saveoptions "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

type Config struct {
	separator                    *byte
	separatorString              *string
	encoding                     *string
	quoteType                    *string
	formatStrategy               *string
	trimLeadingBlankRowAndColumn *bool
	trimTailingBlankCells        *bool
	keepSeparatorsForBlankRow    *bool
	exportArea                   *asposecells.CellArea
	exportQuotePrefix            *bool
	exportAllSheets              *bool
	saveoptions.CommonConfig
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
	engine.LockEngine()
	defer engine.UnlockEngine()
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
		value, err := enums.EncodingType(*c.encoding)
		if err != nil {
			return nil, err
		}
		if err := opts.SetEncoding(value); err != nil {
			return nil, err
		}
	}
	if c.quoteType != nil {
		value, err := enums.TxtValueQuoteType(*c.quoteType)
		if err != nil {
			return nil, err
		}
		if err := opts.SetQuoteType(value); err != nil {
			return nil, err
		}
	}
	if c.formatStrategy != nil {
		value, err := enums.CellValueFormatStrategy(*c.formatStrategy)
		if err != nil {
			return nil, err
		}
		if err := opts.SetFormatStrategy(value); err != nil {
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
	if err := c.ApplyCommon(opts); err != nil {
		return nil, err
	}
	workbook, err := engine.OpenWorkbook(source)
	if err != nil {
		return nil, err
	}
	defer engine.CloseWorkbook(workbook)
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

// WithEncoding sets the encoding: "ascii", "default", "unicode", or "utf8".
func WithEncoding(value string) Option {
	return func(c *Config) {
		c.encoding = &value
	}
}

// WithQuoteType sets the quote type: "always", "minimum", "never", or
// "normal".
func WithQuoteType(value string) Option {
	return func(c *Config) {
		c.quoteType = &value
	}
}

// WithFormatStrategy sets the format strategy: "cellStyle", "displayString",
// "displayStyle", or "none".
func WithFormatStrategy(value string) Option {
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

// Disabled: this option names an engine type, which the toolkit's public API
// must not do — a caller who set it would be tied to the binding
// (docs/design.md §11). Restore it by taking the toolkit-native value instead,
// the way json.WithExportArea takes an "A1:C3" string. The original
// declaration follows verbatim.
/*
func WithExportArea(value *asposecells.CellArea) Option {
	return func(c *Config) {
		c.exportArea = value
	}
}
*/

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
