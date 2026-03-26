package json

import (
	. "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
	"strconv"
)

type Config struct {
	exportStylePool           string
	exportHyperlinkType       string
	skipEmptyRows             string
	sheetIndexes              []int32
	schemas                   []string
	exportArea                *asposecells.CellArea
	hasHeaderRow              string
	exportAsString            string
	indent                    string
	exportNestedStructure     string
	exportEmptyCells          string
	alwaysExportAsJsonObject  string
	toExcelStruct             string
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
	opts, _ := asposecells.NewJsonSaveOptions()

	if len(c.exportStylePool) > 0 {
		if v, err := strconv.ParseBool(c.exportStylePool); err == nil {
			opts.SetExportStylePool(v)
		}
	}
	if v, err := strconv.ParseInt(c.exportHyperlinkType, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToJsonExportHyperlinkType(int32(v)); err2 == nil {
			opts.SetExportHyperlinkType(vv)
		}
	}

	if len(c.skipEmptyRows) > 0 {
		if v, err := strconv.ParseBool(c.skipEmptyRows); err == nil {
			opts.SetSkipEmptyRows(v)
		}
	}
	if c.sheetIndexes != nil {
		opts.SetSheetIndexes(c.sheetIndexes)
	}

	if c.schemas != nil {
		opts.SetSchemas(c.schemas)
	}

	if c.exportArea != nil {
		opts.SetExportArea(c.exportArea)
	}

	if len(c.hasHeaderRow) > 0 {
		if v, err := strconv.ParseBool(c.hasHeaderRow); err == nil {
			opts.SetHasHeaderRow(v)
		}
	}
	if len(c.exportAsString) > 0 {
		if v, err := strconv.ParseBool(c.exportAsString); err == nil {
			opts.SetExportAsString(v)
		}
	}
	if len(c.indent) > 0 {
		opts.SetIndent(c.indent)
	}
	if len(c.exportNestedStructure) > 0 {
		if v, err := strconv.ParseBool(c.exportNestedStructure); err == nil {
			opts.SetExportNestedStructure(v)
		}
	}
	if len(c.exportEmptyCells) > 0 {
		if v, err := strconv.ParseBool(c.exportEmptyCells); err == nil {
			opts.SetExportEmptyCells(v)
		}
	}
	if len(c.alwaysExportAsJsonObject) > 0 {
		if v, err := strconv.ParseBool(c.alwaysExportAsJsonObject); err == nil {
			opts.SetAlwaysExportAsJsonObject(v)
		}
	}
	if len(c.toExcelStruct) > 0 {
		if v, err := strconv.ParseBool(c.toExcelStruct); err == nil {
			opts.SetToExcelStruct(v)
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

func WithExportStylePool(value bool) Option {
	return func(c *Config) {
		c.exportStylePool = strconv.FormatBool(value)
	}
}

func WithExportHyperlinkType(value asposecells.JsonExportHyperlinkType) Option {
	return func(c *Config) {
		c.exportHyperlinkType = strconv.FormatInt(int64(value), 10)
	}
}
func WithSkipEmptyRows(value bool) Option {
	return func(c *Config) {
		c.skipEmptyRows = strconv.FormatBool(value)
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
		c.hasHeaderRow = strconv.FormatBool(value)
	}
}

func WithExportAsString(value bool) Option {
	return func(c *Config) {
		c.exportAsString = strconv.FormatBool(value)
	}
}

func WithIndent(value string) Option {
	return func(c *Config) {
		c.indent = value
	}
}

func WithExportNestedStructure(value bool) Option {
	return func(c *Config) {
		c.exportNestedStructure = strconv.FormatBool(value)
	}
}

func WithExportEmptyCells(value bool) Option {
	return func(c *Config) {
		c.exportEmptyCells = strconv.FormatBool(value)
	}
}

func WithAlwaysExportAsJsonObject(value bool) Option {
	return func(c *Config) {
		c.alwaysExportAsJsonObject = strconv.FormatBool(value)
	}
}

func WithToExcelStruct(value bool) Option {
	return func(c *Config) {
		c.toExcelStruct = strconv.FormatBool(value)
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
