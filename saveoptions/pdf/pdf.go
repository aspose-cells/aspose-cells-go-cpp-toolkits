package pdf

import (
	"strconv"
	"time"

	. "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

type Config struct {
	embedStandardWindowsFonts         string
	bookmark                          *asposecells.PdfBookmarkEntry
	compliance                        string
	securityOptions                   *asposecells.PdfSecurityOptions
	calculateFormula                  string
	pdfCompression                    string
	createdTime                       time.Time
	producer                          string
	optimizationType                  string
	customPropertiesExport            string
	exportDocumentStructure           string
	displayDocTitle                   string
	fontEncoding                      string
	watermark                         *asposecells.RenderingWatermark
	embedAttachments                  string
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

func (c *Config) Apply(source []byte) ([]byte, error) {
	opts, _ := asposecells.NewPdfSaveOptions()

	if len(c.embedStandardWindowsFonts) > 0 {
		if v, err := strconv.ParseBool(c.embedStandardWindowsFonts); err == nil {
			opts.SetEmbedStandardWindowsFonts(v)
		}
	}
	if c.bookmark != nil {
		opts.SetBookmark(c.bookmark)
	}

	if v, err := strconv.ParseInt(c.compliance, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToPdfCompliance(int32(v)); err2 == nil {
			opts.SetCompliance(vv)
		}
	}

	if c.securityOptions != nil {
		opts.SetSecurityOptions(c.securityOptions)
	}

	if len(c.calculateFormula) > 0 {
		if v, err := strconv.ParseBool(c.calculateFormula); err == nil {
			opts.SetCalculateFormula(v)
		}
	}

	if len(c.producer) > 0 {
		opts.SetProducer(c.producer)
	}
	if v, err := strconv.ParseInt(c.optimizationType, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToPdfOptimizationType(int32(v)); err2 == nil {
			opts.SetOptimizationType(vv)
		}
	}

	if v, err := strconv.ParseInt(c.customPropertiesExport, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToPdfCustomPropertiesExport(int32(v)); err2 == nil {
			opts.SetCustomPropertiesExport(vv)
		}
	}

	if len(c.exportDocumentStructure) > 0 {
		if v, err := strconv.ParseBool(c.exportDocumentStructure); err == nil {
			opts.SetExportDocumentStructure(v)
		}
	}
	if len(c.displayDocTitle) > 0 {
		if v, err := strconv.ParseBool(c.displayDocTitle); err == nil {
			opts.SetDisplayDocTitle(v)
		}
	}
	if v, err := strconv.ParseInt(c.fontEncoding, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToPdfFontEncoding(int32(v)); err2 == nil {
			opts.SetFontEncoding(vv)
		}
	}

	if c.watermark != nil {
		opts.SetWatermark(c.watermark)
	}

	if len(c.embedAttachments) > 0 {
		if v, err := strconv.ParseBool(c.embedAttachments); err == nil {
			opts.SetEmbedAttachments(v)
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

func WithEmbedStandardWindowsFonts(value bool) Option {
	return func(c *Config) {
		c.embedStandardWindowsFonts = strconv.FormatBool(value)
	}
}

func WithBookmark(value *asposecells.PdfBookmarkEntry) Option {
	return func(c *Config) {
		c.bookmark = value
	}
}
func WithCompliance(value asposecells.PdfCompliance) Option {
	return func(c *Config) {
		c.compliance = strconv.FormatInt(int64(value), 10)
	}
}
func WithSecurityOptions(value *asposecells.PdfSecurityOptions) Option {
	return func(c *Config) {
		c.securityOptions = value
	}
}
func WithCalculateFormula(value bool) Option {
	return func(c *Config) {
		c.calculateFormula = strconv.FormatBool(value)
	}
}

func WithPdfCompression(value asposecells.PdfCompressionCore) Option {
	return func(c *Config) {
		c.pdfCompression = strconv.FormatInt(int64(value), 10)
	}
}
func WithCreatedTime(value time.Time) Option {
	return func(c *Config) {
		c.createdTime = value
	}
}
func WithProducer(value string) Option {
	return func(c *Config) {
		c.producer = value
	}
}

func WithOptimizationType(value asposecells.PdfOptimizationType) Option {
	return func(c *Config) {
		c.optimizationType = strconv.FormatInt(int64(value), 10)
	}
}
func WithCustomPropertiesExport(value asposecells.PdfCustomPropertiesExport) Option {
	return func(c *Config) {
		c.customPropertiesExport = strconv.FormatInt(int64(value), 10)
	}
}
func WithExportDocumentStructure(value bool) Option {
	return func(c *Config) {
		c.exportDocumentStructure = strconv.FormatBool(value)
	}
}

func WithDisplayDocTitle(value bool) Option {
	return func(c *Config) {
		c.displayDocTitle = strconv.FormatBool(value)
	}
}

func WithFontEncoding(value asposecells.PdfFontEncoding) Option {
	return func(c *Config) {
		c.fontEncoding = strconv.FormatInt(int64(value), 10)
	}
}
func WithWatermark(value *asposecells.RenderingWatermark) Option {
	return func(c *Config) {
		c.watermark = value
	}
}
func WithEmbedAttachments(value bool) Option {
	return func(c *Config) {
		c.embedAttachments = strconv.FormatBool(value)
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
