package txt

import (
	. "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
	"strconv"
)

type Config struct {
	separator                    string
	separatorString              string
	encoding                     string
	quoteType                    string
	formatStrategy               string
	trimLeadingBlankRowAndColumn string
	trimTailingBlankCells        string
	keepSeparatorsForBlankRow    string
	exportArea                   *asposecells.CellArea
	exportQuotePrefix            string
	exportAllSheets              string
	clearData                    string
	cachedFileFolder             string
	validateMergedAreas          string
	mergeAreas                   string
	createDirectory              string
	sortNames                    string
	sortExternalNames            string
	refreshChartCache            string
	checkExcelRestriction        string
	updateSmartArt               string
	encryptDocumentProperties    string
}

// Apply processes the given source byte slice as a Txt file and returns the converted output.
// This method satisfies the saveoptions.SaveOption (or equivalent) interface, enabling Txt-specific export logic.
//
// Parameters:
// - source: A byte slice representing the input spreadsheet or data source. The implementation may interpret
// this as an intermediate format (e.g., XLSX or CSV bytes) and convert it into Txt format.
//
// Returns:
// - []byte: The resulting Txt file content as a byte slice.
// - error: error information.
func (c *Config) Apply(source []byte) ([]byte, error) {
	opts, _ := asposecells.NewTxtSaveOptions()

	if len(c.separator) > 0 {
		if v, err := strconv.ParseInt(c.separator, 10, 64); err == nil {
			opts.SetSeparator(byte(v))
		}
	}
	if len(c.separatorString) > 0 {
		opts.SetSeparatorString(c.separatorString)
	}
	if v, err := strconv.ParseInt(c.encoding, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToEncodingType(int32(v)); err2 == nil {
			opts.SetEncoding(vv)
		}
	}

	if v, err := strconv.ParseInt(c.quoteType, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToTxtValueQuoteType(int32(v)); err2 == nil {
			opts.SetQuoteType(vv)
		}
	}

	if v, err := strconv.ParseInt(c.formatStrategy, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToCellValueFormatStrategy(int32(v)); err2 == nil {
			opts.SetFormatStrategy(vv)
		}
	}

	if len(c.trimLeadingBlankRowAndColumn) > 0 {
		if v, err := strconv.ParseBool(c.trimLeadingBlankRowAndColumn); err == nil {
			opts.SetTrimLeadingBlankRowAndColumn(v)
		}
	}
	if len(c.trimTailingBlankCells) > 0 {
		if v, err := strconv.ParseBool(c.trimTailingBlankCells); err == nil {
			opts.SetTrimTailingBlankCells(v)
		}
	}
	if len(c.keepSeparatorsForBlankRow) > 0 {
		if v, err := strconv.ParseBool(c.keepSeparatorsForBlankRow); err == nil {
			opts.SetKeepSeparatorsForBlankRow(v)
		}
	}
	if c.exportArea != nil {
		opts.SetExportArea(c.exportArea)
	}

	if len(c.exportQuotePrefix) > 0 {
		if v, err := strconv.ParseBool(c.exportQuotePrefix); err == nil {
			opts.SetExportQuotePrefix(v)
		}
	}
	if len(c.exportAllSheets) > 0 {
		if v, err := strconv.ParseBool(c.exportAllSheets); err == nil {
			opts.SetExportAllSheets(v)
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

type Option func(*Config)

func New(opts ...Option) SaveOption {

	cfg := &Config{}

	for _, o := range opts {
		o(cfg)
	}

	return cfg
}

func WithSeparator(value byte) Option {
	return func(c *Config) {
		c.separator = strconv.Itoa(int(value))
	}
}

func WithSeparatorString(value string) Option {
	return func(c *Config) {
		c.separatorString = value
	}
}

func WithEncoding(value asposecells.EncodingType) Option {
	return func(c *Config) {
		c.encoding = strconv.FormatInt(int64(value), 10)
	}
}
func WithQuoteType(value asposecells.TxtValueQuoteType) Option {
	return func(c *Config) {
		c.quoteType = strconv.FormatInt(int64(value), 10)
	}
}
func WithFormatStrategy(value asposecells.CellValueFormatStrategy) Option {
	return func(c *Config) {
		c.formatStrategy = strconv.FormatInt(int64(value), 10)
	}
}
func WithTrimLeadingBlankRowAndColumn(value bool) Option {
	return func(c *Config) {
		c.trimLeadingBlankRowAndColumn = strconv.FormatBool(value)
	}
}

func WithTrimTailingBlankCells(value bool) Option {
	return func(c *Config) {
		c.trimTailingBlankCells = strconv.FormatBool(value)
	}
}

func WithKeepSeparatorsForBlankRow(value bool) Option {
	return func(c *Config) {
		c.keepSeparatorsForBlankRow = strconv.FormatBool(value)
	}
}

func WithExportArea(value *asposecells.CellArea) Option {
	return func(c *Config) {
		c.exportArea = value
	}
}
func WithExportQuotePrefix(value bool) Option {
	return func(c *Config) {
		c.exportQuotePrefix = strconv.FormatBool(value)
	}
}

func WithExportAllSheets(value bool) Option {
	return func(c *Config) {
		c.exportAllSheets = strconv.FormatBool(value)
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
