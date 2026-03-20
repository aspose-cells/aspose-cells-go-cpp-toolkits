package ebook

import (
	. "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
	"strconv"
)

type Config struct {
	ignoreInvisibleShapes            string
	pageTitle                        string
	attachedFilesDirectory           string
	attachedFilesUrlPrefix           string
	defaultFontName                  string
	addGenericFont                   string
	worksheetScalable                string
	isExportComments                 string
	exportCommentsType               string
	disableDownlevelRevealedComments string
	isExpImageToTempDir              string
	imageScalable                    string
	widthScalable                    string
	exportSingleTab                  string
	exportImagesAsBase64             string
	exportActiveWorksheetOnly        string
	exportPrintAreaOnly              string
	exportArea                       *asposecells.CellArea
	parseHtmlTagInCell               string
	htmlCrossStringType              string
	hiddenColDisplayType             string
	hiddenRowDisplayType             string
	encoding                         string
	saveAsSingleFile                 string
	showAllSheets                    string
	exportPageHeaders                string
	exportPageFooters                string
	exportHiddenWorksheet            string
	presentationPreference           string
	cellCssPrefix                    string
	tableCssId                       string
	isFullPathLink                   string
	exportWorksheetCSSSeparately     string
	exportSimilarBorderStyle         string
	mergeEmptyTdType                 string
	exportCellCoordinate             string
	exportExtraHeadings              string
	exportRowColumnHeadings          string
	exportFormula                    string
	addTooltipText                   string
	exportGridLines                  string
	exportBogusRowData               string
	excludeUnusedStyles              string
	exportDocumentProperties         string
	exportWorksheetProperties        string
	exportWorkbookProperties         string
	exportFrameScriptsAndProperties  string
	exportDataOptions                string
	linkTargetType                   string
	isIECompatible                   string
	formatDataIgnoreColumnWidth      string
	calculateFormula                 string
	isJsBrowserCompatible            string
	isMobileCompatible               string
	cssStyles                        string
	hideOverflowWrappedText          string
	isBorderCollapsed                string
	encodeEntityAsCode               string
	officeMathOutputMode             string
	cellNameAttribute                string
	disableCss                       string
	enableCssCustomProperties        string
	htmlVersion                      string
	sheetSet                         *asposecells.SheetSet
	layoutMode                       string
	embeddedFontType                 string
	exportNamedRangeAnchors          string
	dataBarRenderMode                string
	clearData                        string
	cachedFileFolder                 string
	validateMergedAreas              string
	mergeAreas                       string
	createDirectory                  string
	sortNames                        string
	sortExternalNames                string
	refreshChartCache                string
	checkExcelRestriction            string
	updateSmartArt                   string
	encryptDocumentProperties        string
}

func (c *Config) Apply(source []byte) ([]byte, error) {
	opts, _ := asposecells.NewEbookSaveOptions()

	if len(c.ignoreInvisibleShapes) > 0 {
		if v, err := strconv.ParseBool(c.ignoreInvisibleShapes); err == nil {
			opts.SetIgnoreInvisibleShapes(v)
		}
	}
	if len(c.pageTitle) > 0 {
		opts.SetPageTitle(c.pageTitle)
	}
	if len(c.attachedFilesDirectory) > 0 {
		opts.SetAttachedFilesDirectory(c.attachedFilesDirectory)
	}
	if len(c.attachedFilesUrlPrefix) > 0 {
		opts.SetAttachedFilesUrlPrefix(c.attachedFilesUrlPrefix)
	}
	if len(c.defaultFontName) > 0 {
		opts.SetDefaultFontName(c.defaultFontName)
	}
	if len(c.addGenericFont) > 0 {
		if v, err := strconv.ParseBool(c.addGenericFont); err == nil {
			opts.SetAddGenericFont(v)
		}
	}
	if len(c.worksheetScalable) > 0 {
		if v, err := strconv.ParseBool(c.worksheetScalable); err == nil {
			opts.SetWorksheetScalable(v)
		}
	}
	if len(c.isExportComments) > 0 {
		if v, err := strconv.ParseBool(c.isExportComments); err == nil {
			opts.SetIsExportComments(v)
		}
	}
	if v, err := strconv.ParseInt(c.exportCommentsType, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToPrintCommentsType(int32(v)); err2 == nil {
			opts.SetExportCommentsType(vv)
		}
	}

	if len(c.disableDownlevelRevealedComments) > 0 {
		if v, err := strconv.ParseBool(c.disableDownlevelRevealedComments); err == nil {
			opts.SetDisableDownlevelRevealedComments(v)
		}
	}
	if len(c.isExpImageToTempDir) > 0 {
		if v, err := strconv.ParseBool(c.isExpImageToTempDir); err == nil {
			opts.SetIsExpImageToTempDir(v)
		}
	}
	if len(c.imageScalable) > 0 {
		if v, err := strconv.ParseBool(c.imageScalable); err == nil {
			opts.SetImageScalable(v)
		}
	}
	if len(c.widthScalable) > 0 {
		if v, err := strconv.ParseBool(c.widthScalable); err == nil {
			opts.SetWidthScalable(v)
		}
	}
	if len(c.exportSingleTab) > 0 {
		if v, err := strconv.ParseBool(c.exportSingleTab); err == nil {
			opts.SetExportSingleTab(v)
		}
	}
	if len(c.exportImagesAsBase64) > 0 {
		if v, err := strconv.ParseBool(c.exportImagesAsBase64); err == nil {
			opts.SetExportImagesAsBase64(v)
		}
	}
	if len(c.exportActiveWorksheetOnly) > 0 {
		if v, err := strconv.ParseBool(c.exportActiveWorksheetOnly); err == nil {
			opts.SetExportActiveWorksheetOnly(v)
		}
	}
	if len(c.exportPrintAreaOnly) > 0 {
		if v, err := strconv.ParseBool(c.exportPrintAreaOnly); err == nil {
			opts.SetExportPrintAreaOnly(v)
		}
	}
	if c.exportArea != nil {
		opts.SetExportArea(c.exportArea)
	}

	if len(c.parseHtmlTagInCell) > 0 {
		if v, err := strconv.ParseBool(c.parseHtmlTagInCell); err == nil {
			opts.SetParseHtmlTagInCell(v)
		}
	}
	if v, err := strconv.ParseInt(c.htmlCrossStringType, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToHtmlCrossType(int32(v)); err2 == nil {
			opts.SetHtmlCrossStringType(vv)
		}
	}

	if v, err := strconv.ParseInt(c.hiddenColDisplayType, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToHtmlHiddenColDisplayType(int32(v)); err2 == nil {
			opts.SetHiddenColDisplayType(vv)
		}
	}

	if v, err := strconv.ParseInt(c.hiddenRowDisplayType, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToHtmlHiddenRowDisplayType(int32(v)); err2 == nil {
			opts.SetHiddenRowDisplayType(vv)
		}
	}

	if v, err := strconv.ParseInt(c.encoding, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToEncodingType(int32(v)); err2 == nil {
			opts.SetEncoding(vv)
		}
	}

	if len(c.saveAsSingleFile) > 0 {
		if v, err := strconv.ParseBool(c.saveAsSingleFile); err == nil {
			opts.SetSaveAsSingleFile(v)
		}
	}
	if len(c.showAllSheets) > 0 {
		if v, err := strconv.ParseBool(c.showAllSheets); err == nil {
			opts.SetShowAllSheets(v)
		}
	}
	if len(c.exportPageHeaders) > 0 {
		if v, err := strconv.ParseBool(c.exportPageHeaders); err == nil {
			opts.SetExportPageHeaders(v)
		}
	}
	if len(c.exportPageFooters) > 0 {
		if v, err := strconv.ParseBool(c.exportPageFooters); err == nil {
			opts.SetExportPageFooters(v)
		}
	}
	if len(c.exportHiddenWorksheet) > 0 {
		if v, err := strconv.ParseBool(c.exportHiddenWorksheet); err == nil {
			opts.SetExportHiddenWorksheet(v)
		}
	}
	if len(c.presentationPreference) > 0 {
		if v, err := strconv.ParseBool(c.presentationPreference); err == nil {
			opts.SetPresentationPreference(v)
		}
	}
	if len(c.cellCssPrefix) > 0 {
		opts.SetCellCssPrefix(c.cellCssPrefix)
	}
	if len(c.tableCssId) > 0 {
		opts.SetTableCssId(c.tableCssId)
	}
	if len(c.isFullPathLink) > 0 {
		if v, err := strconv.ParseBool(c.isFullPathLink); err == nil {
			opts.SetIsFullPathLink(v)
		}
	}
	if len(c.exportWorksheetCSSSeparately) > 0 {
		if v, err := strconv.ParseBool(c.exportWorksheetCSSSeparately); err == nil {
			opts.SetExportWorksheetCSSSeparately(v)
		}
	}
	if len(c.exportSimilarBorderStyle) > 0 {
		if v, err := strconv.ParseBool(c.exportSimilarBorderStyle); err == nil {
			opts.SetExportSimilarBorderStyle(v)
		}
	}
	if v, err := strconv.ParseInt(c.mergeEmptyTdType, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToMergeEmptyTdType(int32(v)); err2 == nil {
			opts.SetMergeEmptyTdType(vv)
		}
	}

	if len(c.exportCellCoordinate) > 0 {
		if v, err := strconv.ParseBool(c.exportCellCoordinate); err == nil {
			opts.SetExportCellCoordinate(v)
		}
	}
	if len(c.exportExtraHeadings) > 0 {
		if v, err := strconv.ParseBool(c.exportExtraHeadings); err == nil {
			opts.SetExportExtraHeadings(v)
		}
	}
	if len(c.exportRowColumnHeadings) > 0 {
		if v, err := strconv.ParseBool(c.exportRowColumnHeadings); err == nil {
			opts.SetExportRowColumnHeadings(v)
		}
	}
	if len(c.exportFormula) > 0 {
		if v, err := strconv.ParseBool(c.exportFormula); err == nil {
			opts.SetExportFormula(v)
		}
	}
	if len(c.addTooltipText) > 0 {
		if v, err := strconv.ParseBool(c.addTooltipText); err == nil {
			opts.SetAddTooltipText(v)
		}
	}
	if len(c.exportGridLines) > 0 {
		if v, err := strconv.ParseBool(c.exportGridLines); err == nil {
			opts.SetExportGridLines(v)
		}
	}
	if len(c.exportBogusRowData) > 0 {
		if v, err := strconv.ParseBool(c.exportBogusRowData); err == nil {
			opts.SetExportBogusRowData(v)
		}
	}
	if len(c.excludeUnusedStyles) > 0 {
		if v, err := strconv.ParseBool(c.excludeUnusedStyles); err == nil {
			opts.SetExcludeUnusedStyles(v)
		}
	}
	if len(c.exportDocumentProperties) > 0 {
		if v, err := strconv.ParseBool(c.exportDocumentProperties); err == nil {
			opts.SetExportDocumentProperties(v)
		}
	}
	if len(c.exportWorksheetProperties) > 0 {
		if v, err := strconv.ParseBool(c.exportWorksheetProperties); err == nil {
			opts.SetExportWorksheetProperties(v)
		}
	}
	if len(c.exportWorkbookProperties) > 0 {
		if v, err := strconv.ParseBool(c.exportWorkbookProperties); err == nil {
			opts.SetExportWorkbookProperties(v)
		}
	}
	if len(c.exportFrameScriptsAndProperties) > 0 {
		if v, err := strconv.ParseBool(c.exportFrameScriptsAndProperties); err == nil {
			opts.SetExportFrameScriptsAndProperties(v)
		}
	}
	if v, err := strconv.ParseInt(c.exportDataOptions, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToHtmlExportDataOptions(int32(v)); err2 == nil {
			opts.SetExportDataOptions(vv)
		}
	}

	if v, err := strconv.ParseInt(c.linkTargetType, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToHtmlLinkTargetType(int32(v)); err2 == nil {
			opts.SetLinkTargetType(vv)
		}
	}

	if len(c.isIECompatible) > 0 {
		if v, err := strconv.ParseBool(c.isIECompatible); err == nil {
			opts.SetIsIECompatible(v)
		}
	}
	if len(c.formatDataIgnoreColumnWidth) > 0 {
		if v, err := strconv.ParseBool(c.formatDataIgnoreColumnWidth); err == nil {
			opts.SetFormatDataIgnoreColumnWidth(v)
		}
	}
	if len(c.calculateFormula) > 0 {
		if v, err := strconv.ParseBool(c.calculateFormula); err == nil {
			opts.SetCalculateFormula(v)
		}
	}
	if len(c.isJsBrowserCompatible) > 0 {
		if v, err := strconv.ParseBool(c.isJsBrowserCompatible); err == nil {
			opts.SetIsJsBrowserCompatible(v)
		}
	}
	if len(c.isMobileCompatible) > 0 {
		if v, err := strconv.ParseBool(c.isMobileCompatible); err == nil {
			opts.SetIsMobileCompatible(v)
		}
	}
	if len(c.cssStyles) > 0 {
		opts.SetCssStyles(c.cssStyles)
	}
	if len(c.hideOverflowWrappedText) > 0 {
		if v, err := strconv.ParseBool(c.hideOverflowWrappedText); err == nil {
			opts.SetHideOverflowWrappedText(v)
		}
	}
	if len(c.isBorderCollapsed) > 0 {
		if v, err := strconv.ParseBool(c.isBorderCollapsed); err == nil {
			opts.SetIsBorderCollapsed(v)
		}
	}
	if len(c.encodeEntityAsCode) > 0 {
		if v, err := strconv.ParseBool(c.encodeEntityAsCode); err == nil {
			opts.SetEncodeEntityAsCode(v)
		}
	}
	if v, err := strconv.ParseInt(c.officeMathOutputMode, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToHtmlOfficeMathOutputType(int32(v)); err2 == nil {
			opts.SetOfficeMathOutputMode(vv)
		}
	}

	if len(c.cellNameAttribute) > 0 {
		opts.SetCellNameAttribute(c.cellNameAttribute)
	}
	if len(c.disableCss) > 0 {
		if v, err := strconv.ParseBool(c.disableCss); err == nil {
			opts.SetDisableCss(v)
		}
	}
	if len(c.enableCssCustomProperties) > 0 {
		if v, err := strconv.ParseBool(c.enableCssCustomProperties); err == nil {
			opts.SetEnableCssCustomProperties(v)
		}
	}
	if v, err := strconv.ParseInt(c.htmlVersion, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToHtmlVersion(int32(v)); err2 == nil {
			opts.SetHtmlVersion(vv)
		}
	}

	if c.sheetSet != nil {
		opts.SetSheetSet(c.sheetSet)
	}

	if v, err := strconv.ParseInt(c.layoutMode, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToHtmlLayoutMode(int32(v)); err2 == nil {
			opts.SetLayoutMode(vv)
		}
	}

	if v, err := strconv.ParseInt(c.embeddedFontType, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToHtmlEmbeddedFontType(int32(v)); err2 == nil {
			opts.SetEmbeddedFontType(vv)
		}
	}

	if len(c.exportNamedRangeAnchors) > 0 {
		if v, err := strconv.ParseBool(c.exportNamedRangeAnchors); err == nil {
			opts.SetExportNamedRangeAnchors(v)
		}
	}
	if v, err := strconv.ParseInt(c.dataBarRenderMode, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToDataBarRenderMode(int32(v)); err2 == nil {
			opts.SetDataBarRenderMode(vv)
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

func WithIgnoreInvisibleShapes(value bool) Option {
	return func(c *Config) {
		c.ignoreInvisibleShapes = strconv.FormatBool(value)
	}
}

func WithPageTitle(value string) Option {
	return func(c *Config) {
		c.pageTitle = value
	}
}

func WithAttachedFilesDirectory(value string) Option {
	return func(c *Config) {
		c.attachedFilesDirectory = value
	}
}

func WithAttachedFilesUrlPrefix(value string) Option {
	return func(c *Config) {
		c.attachedFilesUrlPrefix = value
	}
}

func WithDefaultFontName(value string) Option {
	return func(c *Config) {
		c.defaultFontName = value
	}
}

func WithAddGenericFont(value bool) Option {
	return func(c *Config) {
		c.addGenericFont = strconv.FormatBool(value)
	}
}

func WithWorksheetScalable(value bool) Option {
	return func(c *Config) {
		c.worksheetScalable = strconv.FormatBool(value)
	}
}

func WithIsExportComments(value bool) Option {
	return func(c *Config) {
		c.isExportComments = strconv.FormatBool(value)
	}
}

func WithExportCommentsType(value asposecells.PrintCommentsType) Option {
	return func(c *Config) {
		c.exportCommentsType = strconv.FormatInt(int64(value), 10)
	}
}
func WithDisableDownlevelRevealedComments(value bool) Option {
	return func(c *Config) {
		c.disableDownlevelRevealedComments = strconv.FormatBool(value)
	}
}

func WithIsExpImageToTempDir(value bool) Option {
	return func(c *Config) {
		c.isExpImageToTempDir = strconv.FormatBool(value)
	}
}

func WithImageScalable(value bool) Option {
	return func(c *Config) {
		c.imageScalable = strconv.FormatBool(value)
	}
}

func WithWidthScalable(value bool) Option {
	return func(c *Config) {
		c.widthScalable = strconv.FormatBool(value)
	}
}

func WithExportSingleTab(value bool) Option {
	return func(c *Config) {
		c.exportSingleTab = strconv.FormatBool(value)
	}
}

func WithExportImagesAsBase64(value bool) Option {
	return func(c *Config) {
		c.exportImagesAsBase64 = strconv.FormatBool(value)
	}
}

func WithExportActiveWorksheetOnly(value bool) Option {
	return func(c *Config) {
		c.exportActiveWorksheetOnly = strconv.FormatBool(value)
	}
}

func WithExportPrintAreaOnly(value bool) Option {
	return func(c *Config) {
		c.exportPrintAreaOnly = strconv.FormatBool(value)
	}
}

func WithExportArea(value *asposecells.CellArea) Option {
	return func(c *Config) {
		c.exportArea = value
	}
}
func WithParseHtmlTagInCell(value bool) Option {
	return func(c *Config) {
		c.parseHtmlTagInCell = strconv.FormatBool(value)
	}
}

func WithHtmlCrossStringType(value asposecells.HtmlCrossType) Option {
	return func(c *Config) {
		c.htmlCrossStringType = strconv.FormatInt(int64(value), 10)
	}
}
func WithHiddenColDisplayType(value asposecells.HtmlHiddenColDisplayType) Option {
	return func(c *Config) {
		c.hiddenColDisplayType = strconv.FormatInt(int64(value), 10)
	}
}
func WithHiddenRowDisplayType(value asposecells.HtmlHiddenRowDisplayType) Option {
	return func(c *Config) {
		c.hiddenRowDisplayType = strconv.FormatInt(int64(value), 10)
	}
}
func WithEncoding(value asposecells.EncodingType) Option {
	return func(c *Config) {
		c.encoding = strconv.FormatInt(int64(value), 10)
	}
}
func WithSaveAsSingleFile(value bool) Option {
	return func(c *Config) {
		c.saveAsSingleFile = strconv.FormatBool(value)
	}
}

func WithShowAllSheets(value bool) Option {
	return func(c *Config) {
		c.showAllSheets = strconv.FormatBool(value)
	}
}

func WithExportPageHeaders(value bool) Option {
	return func(c *Config) {
		c.exportPageHeaders = strconv.FormatBool(value)
	}
}

func WithExportPageFooters(value bool) Option {
	return func(c *Config) {
		c.exportPageFooters = strconv.FormatBool(value)
	}
}

func WithExportHiddenWorksheet(value bool) Option {
	return func(c *Config) {
		c.exportHiddenWorksheet = strconv.FormatBool(value)
	}
}

func WithPresentationPreference(value bool) Option {
	return func(c *Config) {
		c.presentationPreference = strconv.FormatBool(value)
	}
}

func WithCellCssPrefix(value string) Option {
	return func(c *Config) {
		c.cellCssPrefix = value
	}
}

func WithTableCssId(value string) Option {
	return func(c *Config) {
		c.tableCssId = value
	}
}

func WithIsFullPathLink(value bool) Option {
	return func(c *Config) {
		c.isFullPathLink = strconv.FormatBool(value)
	}
}

func WithExportWorksheetCSSSeparately(value bool) Option {
	return func(c *Config) {
		c.exportWorksheetCSSSeparately = strconv.FormatBool(value)
	}
}

func WithExportSimilarBorderStyle(value bool) Option {
	return func(c *Config) {
		c.exportSimilarBorderStyle = strconv.FormatBool(value)
	}
}

func WithMergeEmptyTdType(value asposecells.MergeEmptyTdType) Option {
	return func(c *Config) {
		c.mergeEmptyTdType = strconv.FormatInt(int64(value), 10)
	}
}
func WithExportCellCoordinate(value bool) Option {
	return func(c *Config) {
		c.exportCellCoordinate = strconv.FormatBool(value)
	}
}

func WithExportExtraHeadings(value bool) Option {
	return func(c *Config) {
		c.exportExtraHeadings = strconv.FormatBool(value)
	}
}

func WithExportRowColumnHeadings(value bool) Option {
	return func(c *Config) {
		c.exportRowColumnHeadings = strconv.FormatBool(value)
	}
}

func WithExportFormula(value bool) Option {
	return func(c *Config) {
		c.exportFormula = strconv.FormatBool(value)
	}
}

func WithAddTooltipText(value bool) Option {
	return func(c *Config) {
		c.addTooltipText = strconv.FormatBool(value)
	}
}

func WithExportGridLines(value bool) Option {
	return func(c *Config) {
		c.exportGridLines = strconv.FormatBool(value)
	}
}

func WithExportBogusRowData(value bool) Option {
	return func(c *Config) {
		c.exportBogusRowData = strconv.FormatBool(value)
	}
}

func WithExcludeUnusedStyles(value bool) Option {
	return func(c *Config) {
		c.excludeUnusedStyles = strconv.FormatBool(value)
	}
}

func WithExportDocumentProperties(value bool) Option {
	return func(c *Config) {
		c.exportDocumentProperties = strconv.FormatBool(value)
	}
}

func WithExportWorksheetProperties(value bool) Option {
	return func(c *Config) {
		c.exportWorksheetProperties = strconv.FormatBool(value)
	}
}

func WithExportWorkbookProperties(value bool) Option {
	return func(c *Config) {
		c.exportWorkbookProperties = strconv.FormatBool(value)
	}
}

func WithExportFrameScriptsAndProperties(value bool) Option {
	return func(c *Config) {
		c.exportFrameScriptsAndProperties = strconv.FormatBool(value)
	}
}

func WithExportDataOptions(value asposecells.HtmlExportDataOptions) Option {
	return func(c *Config) {
		c.exportDataOptions = strconv.FormatInt(int64(value), 10)
	}
}
func WithLinkTargetType(value asposecells.HtmlLinkTargetType) Option {
	return func(c *Config) {
		c.linkTargetType = strconv.FormatInt(int64(value), 10)
	}
}
func WithIsIECompatible(value bool) Option {
	return func(c *Config) {
		c.isIECompatible = strconv.FormatBool(value)
	}
}

func WithFormatDataIgnoreColumnWidth(value bool) Option {
	return func(c *Config) {
		c.formatDataIgnoreColumnWidth = strconv.FormatBool(value)
	}
}

func WithCalculateFormula(value bool) Option {
	return func(c *Config) {
		c.calculateFormula = strconv.FormatBool(value)
	}
}

func WithIsJsBrowserCompatible(value bool) Option {
	return func(c *Config) {
		c.isJsBrowserCompatible = strconv.FormatBool(value)
	}
}

func WithIsMobileCompatible(value bool) Option {
	return func(c *Config) {
		c.isMobileCompatible = strconv.FormatBool(value)
	}
}

func WithCssStyles(value string) Option {
	return func(c *Config) {
		c.cssStyles = value
	}
}

func WithHideOverflowWrappedText(value bool) Option {
	return func(c *Config) {
		c.hideOverflowWrappedText = strconv.FormatBool(value)
	}
}

func WithIsBorderCollapsed(value bool) Option {
	return func(c *Config) {
		c.isBorderCollapsed = strconv.FormatBool(value)
	}
}

func WithEncodeEntityAsCode(value bool) Option {
	return func(c *Config) {
		c.encodeEntityAsCode = strconv.FormatBool(value)
	}
}

func WithOfficeMathOutputMode(value asposecells.HtmlOfficeMathOutputType) Option {
	return func(c *Config) {
		c.officeMathOutputMode = strconv.FormatInt(int64(value), 10)
	}
}
func WithCellNameAttribute(value string) Option {
	return func(c *Config) {
		c.cellNameAttribute = value
	}
}

func WithDisableCss(value bool) Option {
	return func(c *Config) {
		c.disableCss = strconv.FormatBool(value)
	}
}

func WithEnableCssCustomProperties(value bool) Option {
	return func(c *Config) {
		c.enableCssCustomProperties = strconv.FormatBool(value)
	}
}

func WithHtmlVersion(value asposecells.HtmlVersion) Option {
	return func(c *Config) {
		c.htmlVersion = strconv.FormatInt(int64(value), 10)
	}
}
func WithSheetSet(value *asposecells.SheetSet) Option {
	return func(c *Config) {
		c.sheetSet = value
	}
}
func WithLayoutMode(value asposecells.HtmlLayoutMode) Option {
	return func(c *Config) {
		c.layoutMode = strconv.FormatInt(int64(value), 10)
	}
}
func WithEmbeddedFontType(value asposecells.HtmlEmbeddedFontType) Option {
	return func(c *Config) {
		c.embeddedFontType = strconv.FormatInt(int64(value), 10)
	}
}
func WithExportNamedRangeAnchors(value bool) Option {
	return func(c *Config) {
		c.exportNamedRangeAnchors = strconv.FormatBool(value)
	}
}

func WithDataBarRenderMode(value asposecells.DataBarRenderMode) Option {
	return func(c *Config) {
		c.dataBarRenderMode = strconv.FormatInt(int64(value), 10)
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
