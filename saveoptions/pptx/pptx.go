package pptx

import (
	. "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
	"strconv"
)

type Config struct {
	ignoreHiddenRows                  string
	saveAsEditableShapes              string
	embedXlsxAsChartDataSource        string
	adjustFontSizeForRowType          string
	exportViewType                    string
	asFlatOpc                         string
	defaultFont                       string
	checkWorkbookDefaultFont          string
	checkFontCompatibility            string
	isFontSubstitutionCharGranularity string
	onePagePerSheet                   string
	allColumnsInOnePagePerSheet       string
	ignoreError                       string
	outputBlankPageWhenNothingToPrint string
	pageIndex                         string
	pageCount                         string
	printingPageType                  string
	gridlineType                      string
	gridlineColor                     *asposecells.Color
	textCrossType                     string
	defaultEditLanguage               string
	sheetSet                          *asposecells.SheetSet
	drawObjectEventHandler            *asposecells.DrawObjectEventHandler
	emfRenderSetting                  string
	customRenderSettings              *asposecells.CustomRenderSettings
	clearData                         string
	cachedFileFolder                  string
	validateMergedAreas               string
	mergeAreas                        string
	createDirectory                   string
	sortNames                         string
	sortExternalNames                 string
	refreshChartCache                 string
	checkExcelRestriction             string
	updateSmartArt                    string
	encryptDocumentProperties         string
}

// Apply processes the given source byte slice as a Pptx file and returns the converted output.
// This method satisfies the saveoptions.SaveOption (or equivalent) interface, enabling Pptx-specific export logic.
//
// Parameters:
// - source: A byte slice representing the input spreadsheet or data source. The implementation may interpret
// this as an intermediate format (e.g., XLSX or CSV bytes) and convert it into Pptx format.
//
// Returns:
// - []byte: The resulting Pptx file content as a byte slice.
// - error: error information.
func (c *Config) Apply(source []byte) ([]byte, error) {
	opts, _ := asposecells.NewPptxSaveOptions()

	if len(c.ignoreHiddenRows) > 0 {
		if v, err := strconv.ParseBool(c.ignoreHiddenRows); err == nil {
			opts.SetIgnoreHiddenRows(v)
		}
	}
	if len(c.saveAsEditableShapes) > 0 {
		if v, err := strconv.ParseBool(c.saveAsEditableShapes); err == nil {
			opts.SetSaveAsEditableShapes(v)
		}
	}
	if len(c.embedXlsxAsChartDataSource) > 0 {
		if v, err := strconv.ParseBool(c.embedXlsxAsChartDataSource); err == nil {
			opts.SetEmbedXlsxAsChartDataSource(v)
		}
	}
	if v, err := strconv.ParseInt(c.adjustFontSizeForRowType, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToAdjustFontSizeForRowType(int32(v)); err2 == nil {
			opts.SetAdjustFontSizeForRowType(vv)
		}
	}

	if v, err := strconv.ParseInt(c.exportViewType, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToSlideViewType(int32(v)); err2 == nil {
			opts.SetExportViewType(vv)
		}
	}

	if len(c.asFlatOpc) > 0 {
		if v, err := strconv.ParseBool(c.asFlatOpc); err == nil {
			opts.SetAsFlatOpc(v)
		}
	}
	if len(c.defaultFont) > 0 {
		opts.SetDefaultFont(c.defaultFont)
	}
	if len(c.checkWorkbookDefaultFont) > 0 {
		if v, err := strconv.ParseBool(c.checkWorkbookDefaultFont); err == nil {
			opts.SetCheckWorkbookDefaultFont(v)
		}
	}
	if len(c.checkFontCompatibility) > 0 {
		if v, err := strconv.ParseBool(c.checkFontCompatibility); err == nil {
			opts.SetCheckFontCompatibility(v)
		}
	}
	if len(c.isFontSubstitutionCharGranularity) > 0 {
		if v, err := strconv.ParseBool(c.isFontSubstitutionCharGranularity); err == nil {
			opts.SetIsFontSubstitutionCharGranularity(v)
		}
	}
	if len(c.onePagePerSheet) > 0 {
		if v, err := strconv.ParseBool(c.onePagePerSheet); err == nil {
			opts.SetOnePagePerSheet(v)
		}
	}
	if len(c.allColumnsInOnePagePerSheet) > 0 {
		if v, err := strconv.ParseBool(c.allColumnsInOnePagePerSheet); err == nil {
			opts.SetAllColumnsInOnePagePerSheet(v)
		}
	}
	if len(c.ignoreError) > 0 {
		if v, err := strconv.ParseBool(c.ignoreError); err == nil {
			opts.SetIgnoreError(v)
		}
	}
	if len(c.outputBlankPageWhenNothingToPrint) > 0 {
		if v, err := strconv.ParseBool(c.outputBlankPageWhenNothingToPrint); err == nil {
			opts.SetOutputBlankPageWhenNothingToPrint(v)
		}
	}
	if len(c.pageIndex) > 0 {
		if v, err := strconv.ParseInt(c.pageIndex, 10, 32); err == nil {
			opts.SetPageIndex(int32(v))
		}
	}
	if len(c.pageCount) > 0 {
		if v, err := strconv.ParseInt(c.pageCount, 10, 32); err == nil {
			opts.SetPageCount(int32(v))
		}
	}
	if v, err := strconv.ParseInt(c.printingPageType, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToPrintingPageType(int32(v)); err2 == nil {
			opts.SetPrintingPageType(vv)
		}
	}

	if v, err := strconv.ParseInt(c.gridlineType, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToGridlineType(int32(v)); err2 == nil {
			opts.SetGridlineType(vv)
		}
	}

	if c.gridlineColor != nil {
		opts.SetGridlineColor(c.gridlineColor)
	}

	if v, err := strconv.ParseInt(c.textCrossType, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToTextCrossType(int32(v)); err2 == nil {
			opts.SetTextCrossType(vv)
		}
	}

	if v, err := strconv.ParseInt(c.defaultEditLanguage, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToDefaultEditLanguage(int32(v)); err2 == nil {
			opts.SetDefaultEditLanguage(vv)
		}
	}

	if c.sheetSet != nil {
		opts.SetSheetSet(c.sheetSet)
	}

	if c.drawObjectEventHandler != nil {
		opts.SetDrawObjectEventHandler(c.drawObjectEventHandler)
	}

	if v, err := strconv.ParseInt(c.emfRenderSetting, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToEmfRenderSetting(int32(v)); err2 == nil {
			opts.SetEmfRenderSetting(vv)
		}
	}

	if c.customRenderSettings != nil {
		opts.SetCustomRenderSettings(c.customRenderSettings)
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

func WithIgnoreHiddenRows(value bool) Option {
	return func(c *Config) {
		c.ignoreHiddenRows = strconv.FormatBool(value)
	}
}

func WithSaveAsEditableShapes(value bool) Option {
	return func(c *Config) {
		c.saveAsEditableShapes = strconv.FormatBool(value)
	}
}

func WithEmbedXlsxAsChartDataSource(value bool) Option {
	return func(c *Config) {
		c.embedXlsxAsChartDataSource = strconv.FormatBool(value)
	}
}

func WithAdjustFontSizeForRowType(value asposecells.AdjustFontSizeForRowType) Option {
	return func(c *Config) {
		c.adjustFontSizeForRowType = strconv.FormatInt(int64(value), 10)
	}
}
func WithExportViewType(value asposecells.SlideViewType) Option {
	return func(c *Config) {
		c.exportViewType = strconv.FormatInt(int64(value), 10)
	}
}
func WithAsFlatOpc(value bool) Option {
	return func(c *Config) {
		c.asFlatOpc = strconv.FormatBool(value)
	}
}

func WithDefaultFont(value string) Option {
	return func(c *Config) {
		c.defaultFont = value
	}
}

func WithCheckWorkbookDefaultFont(value bool) Option {
	return func(c *Config) {
		c.checkWorkbookDefaultFont = strconv.FormatBool(value)
	}
}

func WithCheckFontCompatibility(value bool) Option {
	return func(c *Config) {
		c.checkFontCompatibility = strconv.FormatBool(value)
	}
}

func WithIsFontSubstitutionCharGranularity(value bool) Option {
	return func(c *Config) {
		c.isFontSubstitutionCharGranularity = strconv.FormatBool(value)
	}
}

func WithOnePagePerSheet(value bool) Option {
	return func(c *Config) {
		c.onePagePerSheet = strconv.FormatBool(value)
	}
}

func WithAllColumnsInOnePagePerSheet(value bool) Option {
	return func(c *Config) {
		c.allColumnsInOnePagePerSheet = strconv.FormatBool(value)
	}
}

func WithIgnoreError(value bool) Option {
	return func(c *Config) {
		c.ignoreError = strconv.FormatBool(value)
	}
}

func WithOutputBlankPageWhenNothingToPrint(value bool) Option {
	return func(c *Config) {
		c.outputBlankPageWhenNothingToPrint = strconv.FormatBool(value)
	}
}

func WithPageIndex(value int32) Option {
	return func(c *Config) {
		c.pageIndex = strconv.FormatInt(int64(value), 10)
	}
}

func WithPageCount(value int32) Option {
	return func(c *Config) {
		c.pageCount = strconv.FormatInt(int64(value), 10)
	}
}

func WithPrintingPageType(value asposecells.PrintingPageType) Option {
	return func(c *Config) {
		c.printingPageType = strconv.FormatInt(int64(value), 10)
	}
}
func WithGridlineType(value asposecells.GridlineType) Option {
	return func(c *Config) {
		c.gridlineType = strconv.FormatInt(int64(value), 10)
	}
}
func WithGridlineColor(value *asposecells.Color) Option {
	return func(c *Config) {
		c.gridlineColor = value
	}
}
func WithTextCrossType(value asposecells.TextCrossType) Option {
	return func(c *Config) {
		c.textCrossType = strconv.FormatInt(int64(value), 10)
	}
}
func WithDefaultEditLanguage(value asposecells.DefaultEditLanguage) Option {
	return func(c *Config) {
		c.defaultEditLanguage = strconv.FormatInt(int64(value), 10)
	}
}
func WithSheetSet(value *asposecells.SheetSet) Option {
	return func(c *Config) {
		c.sheetSet = value
	}
}
func WithDrawObjectEventHandler(value *asposecells.DrawObjectEventHandler) Option {
	return func(c *Config) {
		c.drawObjectEventHandler = value
	}
}
func WithEmfRenderSetting(value asposecells.EmfRenderSetting) Option {
	return func(c *Config) {
		c.emfRenderSetting = strconv.FormatInt(int64(value), 10)
	}
}
func WithCustomRenderSettings(value *asposecells.CustomRenderSettings) Option {
	return func(c *Config) {
		c.customRenderSettings = value
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
