package markdown

import (
	"strconv"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/formats"
	. "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

type Config struct {
	encoding                   string
	formatStrategy             string
	lineSeparator              string
	tableHeaderType            string
	sheetSet                   *asposecells.SheetSet
	exportImagesAsBase64       string
	calculateFormula           string
	exportHyperlinkAsReference string
	alignColumnPadding         string
	splitTablesByBlankRow      string
	officeMathOutputType       string
	clearData                  string
	cachedFileFolder           string
	validateMergedAreas        string
	mergeAreas                 string
	createDirectory            string
	sortNames                  string
	sortExternalNames          string
	refreshChartCache          string
	checkExcelRestriction      string
	updateSmartArt             string
	encryptDocumentProperties  string
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
	opts, _ := asposecells.NewMarkdownSaveOptions()

	if v, err := strconv.ParseInt(c.encoding, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToEncodingType(int32(v)); err2 == nil {
			opts.SetEncoding(vv)
		}
	}

	if v, err := strconv.ParseInt(c.formatStrategy, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToCellValueFormatStrategy(int32(v)); err2 == nil {
			opts.SetFormatStrategy(vv)
		}
	}

	if len(c.lineSeparator) > 0 {
		opts.SetLineSeparator(c.lineSeparator)
	}
	if v, err := strconv.ParseInt(c.tableHeaderType, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToMarkdownTableHeaderType(int32(v)); err2 == nil {
			opts.SetTableHeaderType(vv)
		}
	}

	if c.sheetSet != nil {
		opts.SetSheetSet(c.sheetSet)
	}

	if len(c.exportImagesAsBase64) > 0 {
		if v, err := strconv.ParseBool(c.exportImagesAsBase64); err == nil {
			opts.SetExportImagesAsBase64(v)
		}
	}
	if len(c.calculateFormula) > 0 {
		if v, err := strconv.ParseBool(c.calculateFormula); err == nil {
			opts.SetCalculateFormula(v)
		}
	}
	if len(c.exportHyperlinkAsReference) > 0 {
		if v, err := strconv.ParseBool(c.exportHyperlinkAsReference); err == nil {
			opts.SetExportHyperlinkAsReference(v)
		}
	}
	if len(c.alignColumnPadding) > 0 {
		if v, err := strconv.ParseInt(c.alignColumnPadding, 10, 64); err == nil {
			opts.SetAlignColumnPadding(byte(v))
		}
	}
	if len(c.splitTablesByBlankRow) > 0 {
		if v, err := strconv.ParseBool(c.splitTablesByBlankRow); err == nil {
			opts.SetSplitTablesByBlankRow(v)
		}
	}
	//if v, err := strconv.ParseInt(c.officeMathOutputType, 10, 32); err == nil {
	//	if vv, err2 := asposecells.Int32ToHtmlOfficeMathOutputType(int32(v)); err2 == nil {
	//		opts.SetOfficeMathOutputMode(vv)
	//	}
	//}

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
	return "md"
}

type Option func(*Config)

func init() {
	formats.Register("md", func() SaveOption {
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
func New(opts ...Option) SaveOption {

	cfg := &Config{}

	for _, o := range opts {
		o(cfg)
	}

	return cfg
}

func WithEncoding(value asposecells.EncodingType) Option {
	return func(c *Config) {
		c.encoding = strconv.FormatInt(int64(value), 10)
	}
}
func WithFormatStrategy(value asposecells.CellValueFormatStrategy) Option {
	return func(c *Config) {
		c.formatStrategy = strconv.FormatInt(int64(value), 10)
	}
}
func WithLineSeparator(value string) Option {
	return func(c *Config) {
		c.lineSeparator = value
	}
}

func WithTableHeaderType(value asposecells.MarkdownTableHeaderType) Option {
	return func(c *Config) {
		c.tableHeaderType = strconv.FormatInt(int64(value), 10)
	}
}
func WithSheetSet(value *asposecells.SheetSet) Option {
	return func(c *Config) {
		c.sheetSet = value
	}
}
func WithExportImagesAsBase64(value bool) Option {
	return func(c *Config) {
		c.exportImagesAsBase64 = strconv.FormatBool(value)
	}
}

func WithCalculateFormula(value bool) Option {
	return func(c *Config) {
		c.calculateFormula = strconv.FormatBool(value)
	}
}

func WithExportHyperlinkAsReference(value bool) Option {
	return func(c *Config) {
		c.exportHyperlinkAsReference = strconv.FormatBool(value)
	}
}

func WithAlignColumnPadding(value byte) Option {
	return func(c *Config) {
		c.alignColumnPadding = strconv.Itoa(int(value))
	}
}

func WithSplitTablesByBlankRow(value bool) Option {
	return func(c *Config) {
		c.splitTablesByBlankRow = strconv.FormatBool(value)
	}
}

func WithOfficeMathOutputType(value asposecells.HtmlOfficeMathOutputType) Option {
	return func(c *Config) {
		c.officeMathOutputType = strconv.FormatInt(int64(value), 10)
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
