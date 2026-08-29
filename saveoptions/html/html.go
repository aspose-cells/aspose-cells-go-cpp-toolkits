package html

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	saveoptions "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// Config holds the native-typed option values for Html save options.
//
// Pointer fields carry presence semantics: a nil pointer means the option
// was never set (so the native default is kept), while a non-nil pointer
// means the caller explicitly requested the value, including zero/false.
type Config struct {
	ignoreInvisibleShapes            *bool
	pageTitle                        *string
	attachedFilesDirectory           *string
	attachedFilesUrlPrefix           *string
	defaultFontName                  *string
	addGenericFont                   *bool
	worksheetScalable                *bool
	isExportComments                 *bool
	exportCommentsType               *asposecells.PrintCommentsType
	disableDownlevelRevealedComments *bool
	isExpImageToTempDir              *bool
	imageScalable                    *bool
	widthScalable                    *bool
	exportSingleTab                  *bool
	exportImagesAsBase64             *bool
	exportActiveWorksheetOnly        *bool
	exportPrintAreaOnly              *bool
	exportArea                       *asposecells.CellArea
	parseHtmlTagInCell               *bool
	htmlCrossStringType              *asposecells.HtmlCrossType
	hiddenColDisplayType             *asposecells.HtmlHiddenColDisplayType
	hiddenRowDisplayType             *asposecells.HtmlHiddenRowDisplayType
	encoding                         *asposecells.EncodingType
	saveAsSingleFile                 *bool
	showAllSheets                    *bool
	exportPageHeaders                *bool
	exportPageFooters                *bool
	exportHiddenWorksheet            *bool
	presentationPreference           *bool
	cellCssPrefix                    *string
	tableCssId                       *string
	isFullPathLink                   *bool
	exportWorksheetCSSSeparately     *bool
	exportSimilarBorderStyle         *bool
	mergeEmptyTdType                 *asposecells.MergeEmptyTdType
	exportCellCoordinate             *bool
	exportExtraHeadings              *bool
	exportRowColumnHeadings          *bool
	exportFormula                    *bool
	addTooltipText                   *bool
	exportGridLines                  *bool
	exportBogusRowData               *bool
	excludeUnusedStyles              *bool
	exportDocumentProperties         *bool
	exportWorksheetProperties        *bool
	exportWorkbookProperties         *bool
	exportFrameScriptsAndProperties  *bool
	exportDataOptions                *asposecells.HtmlExportDataOptions
	linkTargetType                   *asposecells.HtmlLinkTargetType
	isIECompatible                   *bool
	formatDataIgnoreColumnWidth      *bool
	calculateFormula                 *bool
	isJsBrowserCompatible            *bool
	isMobileCompatible               *bool
	cssStyles                        *string
	hideOverflowWrappedText          *bool
	isBorderCollapsed                *bool
	encodeEntityAsCode               *bool
	officeMathOutputMode             *asposecells.HtmlOfficeMathOutputType
	cellNameAttribute                *string
	disableCss                       *bool
	enableCssCustomProperties        *bool
	htmlVersion                      *asposecells.HtmlVersion
	sheetSet                         *asposecells.SheetSet
	layoutMode                       *asposecells.HtmlLayoutMode
	embeddedFontType                 *asposecells.HtmlEmbeddedFontType
	exportNamedRangeAnchors          *bool
	dataBarRenderMode                *asposecells.DataBarRenderMode
	clearData                        *bool
	cachedFileFolder                 *string
	validateMergedAreas              *bool
	mergeAreas                       *bool
	createDirectory                  *bool
	sortNames                        *bool
	sortExternalNames                *bool
	refreshChartCache                *bool
	checkExcelRestriction            *bool
	updateSmartArt                   *bool
	encryptDocumentProperties        *bool
}

// Apply processes the given source byte slice as a Html file and returns the converted output.
// This method satisfies the saveoptions.SaveOption (or equivalent) interface, enabling Html-specific export logic.
//
// Parameters:
// - source: A byte slice representing the input spreadsheet or data source. The implementation may interpret
// this as an intermediate format (e.g., XLSX or CSV bytes) and convert it into Html format.
//
// Returns:
// - []byte: The resulting Html file content as a byte slice.
// - error: error information.
func (c *Config) Apply(source []byte) ([]byte, error) {
	opts, err := asposecells.NewHtmlSaveOptions()
	if err != nil {
		return nil, err
	}

	if c.ignoreInvisibleShapes != nil {
		if err := opts.SetIgnoreInvisibleShapes(*c.ignoreInvisibleShapes); err != nil {
			return nil, err
		}
	}
	if c.pageTitle != nil {
		if err := opts.SetPageTitle(*c.pageTitle); err != nil {
			return nil, err
		}
	}
	if c.attachedFilesDirectory != nil {
		if err := opts.SetAttachedFilesDirectory(*c.attachedFilesDirectory); err != nil {
			return nil, err
		}
	}
	if c.attachedFilesUrlPrefix != nil {
		if err := opts.SetAttachedFilesUrlPrefix(*c.attachedFilesUrlPrefix); err != nil {
			return nil, err
		}
	}
	if c.defaultFontName != nil {
		if err := opts.SetDefaultFontName(*c.defaultFontName); err != nil {
			return nil, err
		}
	}
	if c.addGenericFont != nil {
		if err := opts.SetAddGenericFont(*c.addGenericFont); err != nil {
			return nil, err
		}
	}
	if c.worksheetScalable != nil {
		if err := opts.SetWorksheetScalable(*c.worksheetScalable); err != nil {
			return nil, err
		}
	}
	if c.isExportComments != nil {
		if err := opts.SetIsExportComments(*c.isExportComments); err != nil {
			return nil, err
		}
	}
	if c.exportCommentsType != nil {
		if err := opts.SetExportCommentsType(*c.exportCommentsType); err != nil {
			return nil, err
		}
	}
	if c.disableDownlevelRevealedComments != nil {
		if err := opts.SetDisableDownlevelRevealedComments(*c.disableDownlevelRevealedComments); err != nil {
			return nil, err
		}
	}
	if c.isExpImageToTempDir != nil {
		if err := opts.SetIsExpImageToTempDir(*c.isExpImageToTempDir); err != nil {
			return nil, err
		}
	}
	if c.imageScalable != nil {
		if err := opts.SetImageScalable(*c.imageScalable); err != nil {
			return nil, err
		}
	}
	if c.widthScalable != nil {
		if err := opts.SetWidthScalable(*c.widthScalable); err != nil {
			return nil, err
		}
	}
	if c.exportSingleTab != nil {
		if err := opts.SetExportSingleTab(*c.exportSingleTab); err != nil {
			return nil, err
		}
	}
	if c.exportImagesAsBase64 != nil {
		if err := opts.SetExportImagesAsBase64(*c.exportImagesAsBase64); err != nil {
			return nil, err
		}
	}
	if c.exportActiveWorksheetOnly != nil {
		if err := opts.SetExportActiveWorksheetOnly(*c.exportActiveWorksheetOnly); err != nil {
			return nil, err
		}
	}
	if c.exportPrintAreaOnly != nil {
		if err := opts.SetExportPrintAreaOnly(*c.exportPrintAreaOnly); err != nil {
			return nil, err
		}
	}
	if c.exportArea != nil {
		if err := opts.SetExportArea(c.exportArea); err != nil {
			return nil, err
		}
	}
	if c.parseHtmlTagInCell != nil {
		if err := opts.SetParseHtmlTagInCell(*c.parseHtmlTagInCell); err != nil {
			return nil, err
		}
	}
	if c.htmlCrossStringType != nil {
		if err := opts.SetHtmlCrossStringType(*c.htmlCrossStringType); err != nil {
			return nil, err
		}
	}
	if c.hiddenColDisplayType != nil {
		if err := opts.SetHiddenColDisplayType(*c.hiddenColDisplayType); err != nil {
			return nil, err
		}
	}
	if c.hiddenRowDisplayType != nil {
		if err := opts.SetHiddenRowDisplayType(*c.hiddenRowDisplayType); err != nil {
			return nil, err
		}
	}
	if c.encoding != nil {
		if err := opts.SetEncoding(*c.encoding); err != nil {
			return nil, err
		}
	}
	if c.saveAsSingleFile != nil {
		if err := opts.SetSaveAsSingleFile(*c.saveAsSingleFile); err != nil {
			return nil, err
		}
	}
	if c.showAllSheets != nil {
		if err := opts.SetShowAllSheets(*c.showAllSheets); err != nil {
			return nil, err
		}
	}
	if c.exportPageHeaders != nil {
		if err := opts.SetExportPageHeaders(*c.exportPageHeaders); err != nil {
			return nil, err
		}
	}
	if c.exportPageFooters != nil {
		if err := opts.SetExportPageFooters(*c.exportPageFooters); err != nil {
			return nil, err
		}
	}
	if c.exportHiddenWorksheet != nil {
		if err := opts.SetExportHiddenWorksheet(*c.exportHiddenWorksheet); err != nil {
			return nil, err
		}
	}
	if c.presentationPreference != nil {
		if err := opts.SetPresentationPreference(*c.presentationPreference); err != nil {
			return nil, err
		}
	}
	if c.cellCssPrefix != nil {
		if err := opts.SetCellCssPrefix(*c.cellCssPrefix); err != nil {
			return nil, err
		}
	}
	if c.tableCssId != nil {
		if err := opts.SetTableCssId(*c.tableCssId); err != nil {
			return nil, err
		}
	}
	if c.isFullPathLink != nil {
		if err := opts.SetIsFullPathLink(*c.isFullPathLink); err != nil {
			return nil, err
		}
	}
	if c.exportWorksheetCSSSeparately != nil {
		if err := opts.SetExportWorksheetCSSSeparately(*c.exportWorksheetCSSSeparately); err != nil {
			return nil, err
		}
	}
	if c.exportSimilarBorderStyle != nil {
		if err := opts.SetExportSimilarBorderStyle(*c.exportSimilarBorderStyle); err != nil {
			return nil, err
		}
	}
	if c.mergeEmptyTdType != nil {
		if err := opts.SetMergeEmptyTdType(*c.mergeEmptyTdType); err != nil {
			return nil, err
		}
	}
	if c.exportCellCoordinate != nil {
		if err := opts.SetExportCellCoordinate(*c.exportCellCoordinate); err != nil {
			return nil, err
		}
	}
	if c.exportExtraHeadings != nil {
		if err := opts.SetExportExtraHeadings(*c.exportExtraHeadings); err != nil {
			return nil, err
		}
	}
	if c.exportRowColumnHeadings != nil {
		if err := opts.SetExportRowColumnHeadings(*c.exportRowColumnHeadings); err != nil {
			return nil, err
		}
	}
	if c.exportFormula != nil {
		if err := opts.SetExportFormula(*c.exportFormula); err != nil {
			return nil, err
		}
	}
	if c.addTooltipText != nil {
		if err := opts.SetAddTooltipText(*c.addTooltipText); err != nil {
			return nil, err
		}
	}
	if c.exportGridLines != nil {
		if err := opts.SetExportGridLines(*c.exportGridLines); err != nil {
			return nil, err
		}
	}
	if c.exportBogusRowData != nil {
		if err := opts.SetExportBogusRowData(*c.exportBogusRowData); err != nil {
			return nil, err
		}
	}
	if c.excludeUnusedStyles != nil {
		if err := opts.SetExcludeUnusedStyles(*c.excludeUnusedStyles); err != nil {
			return nil, err
		}
	}
	if c.exportDocumentProperties != nil {
		if err := opts.SetExportDocumentProperties(*c.exportDocumentProperties); err != nil {
			return nil, err
		}
	}
	if c.exportWorksheetProperties != nil {
		if err := opts.SetExportWorksheetProperties(*c.exportWorksheetProperties); err != nil {
			return nil, err
		}
	}
	if c.exportWorkbookProperties != nil {
		if err := opts.SetExportWorkbookProperties(*c.exportWorkbookProperties); err != nil {
			return nil, err
		}
	}
	if c.exportFrameScriptsAndProperties != nil {
		if err := opts.SetExportFrameScriptsAndProperties(*c.exportFrameScriptsAndProperties); err != nil {
			return nil, err
		}
	}
	if c.exportDataOptions != nil {
		if err := opts.SetExportDataOptions(*c.exportDataOptions); err != nil {
			return nil, err
		}
	}
	if c.linkTargetType != nil {
		if err := opts.SetLinkTargetType(*c.linkTargetType); err != nil {
			return nil, err
		}
	}
	if c.isIECompatible != nil {
		if err := opts.SetIsIECompatible(*c.isIECompatible); err != nil {
			return nil, err
		}
	}
	if c.formatDataIgnoreColumnWidth != nil {
		if err := opts.SetFormatDataIgnoreColumnWidth(*c.formatDataIgnoreColumnWidth); err != nil {
			return nil, err
		}
	}
	if c.calculateFormula != nil {
		if err := opts.SetCalculateFormula(*c.calculateFormula); err != nil {
			return nil, err
		}
	}
	if c.isJsBrowserCompatible != nil {
		if err := opts.SetIsJsBrowserCompatible(*c.isJsBrowserCompatible); err != nil {
			return nil, err
		}
	}
	if c.isMobileCompatible != nil {
		if err := opts.SetIsMobileCompatible(*c.isMobileCompatible); err != nil {
			return nil, err
		}
	}
	if c.cssStyles != nil {
		if err := opts.SetCssStyles(*c.cssStyles); err != nil {
			return nil, err
		}
	}
	if c.hideOverflowWrappedText != nil {
		if err := opts.SetHideOverflowWrappedText(*c.hideOverflowWrappedText); err != nil {
			return nil, err
		}
	}
	if c.isBorderCollapsed != nil {
		if err := opts.SetIsBorderCollapsed(*c.isBorderCollapsed); err != nil {
			return nil, err
		}
	}
	if c.encodeEntityAsCode != nil {
		if err := opts.SetEncodeEntityAsCode(*c.encodeEntityAsCode); err != nil {
			return nil, err
		}
	}
	if c.officeMathOutputMode != nil {
		if err := opts.SetOfficeMathOutputMode(*c.officeMathOutputMode); err != nil {
			return nil, err
		}
	}
	if c.cellNameAttribute != nil {
		if err := opts.SetCellNameAttribute(*c.cellNameAttribute); err != nil {
			return nil, err
		}
	}
	if c.disableCss != nil {
		if err := opts.SetDisableCss(*c.disableCss); err != nil {
			return nil, err
		}
	}
	if c.enableCssCustomProperties != nil {
		if err := opts.SetEnableCssCustomProperties(*c.enableCssCustomProperties); err != nil {
			return nil, err
		}
	}
	if c.htmlVersion != nil {
		if err := opts.SetHtmlVersion(*c.htmlVersion); err != nil {
			return nil, err
		}
	}
	if c.sheetSet != nil {
		if err := opts.SetSheetSet(c.sheetSet); err != nil {
			return nil, err
		}
	}
	if c.layoutMode != nil {
		if err := opts.SetLayoutMode(*c.layoutMode); err != nil {
			return nil, err
		}
	}
	if c.embeddedFontType != nil {
		if err := opts.SetEmbeddedFontType(*c.embeddedFontType); err != nil {
			return nil, err
		}
	}
	if c.exportNamedRangeAnchors != nil {
		if err := opts.SetExportNamedRangeAnchors(*c.exportNamedRangeAnchors); err != nil {
			return nil, err
		}
	}
	if c.dataBarRenderMode != nil {
		if err := opts.SetDataBarRenderMode(*c.dataBarRenderMode); err != nil {
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
	return "html"
}

type Option func(*Config)

func init() {
	formats.Register("html", func() saveoptions.SaveOption {
		return New()
	})
}

// New creates a new instance of html save options
//
// The New function creates an instance of html SaveOption using the Functional Options Pattern. This function accepts a variable number of Option function parameters, and each Option function modifies the configuration of SaveOption.
//
// Parameters:
//
//	opts ... Option - A variable number of option functions used to configure SaveOption
//
// Return value:
// html SaveOption - Configured instance of the saved option
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

func WithIgnoreInvisibleShapes(value bool) Option {
	return func(c *Config) {
		c.ignoreInvisibleShapes = &value
	}
}

func WithPageTitle(value string) Option {
	return func(c *Config) {
		c.pageTitle = &value
	}
}

func WithAttachedFilesDirectory(value string) Option {
	return func(c *Config) {
		c.attachedFilesDirectory = &value
	}
}

func WithAttachedFilesUrlPrefix(value string) Option {
	return func(c *Config) {
		c.attachedFilesUrlPrefix = &value
	}
}

func WithDefaultFontName(value string) Option {
	return func(c *Config) {
		c.defaultFontName = &value
	}
}

func WithAddGenericFont(value bool) Option {
	return func(c *Config) {
		c.addGenericFont = &value
	}
}

func WithWorksheetScalable(value bool) Option {
	return func(c *Config) {
		c.worksheetScalable = &value
	}
}

func WithIsExportComments(value bool) Option {
	return func(c *Config) {
		c.isExportComments = &value
	}
}

func WithExportCommentsType(value asposecells.PrintCommentsType) Option {
	return func(c *Config) {
		c.exportCommentsType = &value
	}
}
func WithDisableDownlevelRevealedComments(value bool) Option {
	return func(c *Config) {
		c.disableDownlevelRevealedComments = &value
	}
}

func WithIsExpImageToTempDir(value bool) Option {
	return func(c *Config) {
		c.isExpImageToTempDir = &value
	}
}

func WithImageScalable(value bool) Option {
	return func(c *Config) {
		c.imageScalable = &value
	}
}

func WithWidthScalable(value bool) Option {
	return func(c *Config) {
		c.widthScalable = &value
	}
}

func WithExportSingleTab(value bool) Option {
	return func(c *Config) {
		c.exportSingleTab = &value
	}
}

func WithExportImagesAsBase64(value bool) Option {
	return func(c *Config) {
		c.exportImagesAsBase64 = &value
	}
}

func WithExportActiveWorksheetOnly(value bool) Option {
	return func(c *Config) {
		c.exportActiveWorksheetOnly = &value
	}
}

func WithExportPrintAreaOnly(value bool) Option {
	return func(c *Config) {
		c.exportPrintAreaOnly = &value
	}
}

func WithExportArea(value *asposecells.CellArea) Option {
	return func(c *Config) {
		c.exportArea = value
	}
}
func WithParseHtmlTagInCell(value bool) Option {
	return func(c *Config) {
		c.parseHtmlTagInCell = &value
	}
}

func WithHtmlCrossStringType(value asposecells.HtmlCrossType) Option {
	return func(c *Config) {
		c.htmlCrossStringType = &value
	}
}
func WithHiddenColDisplayType(value asposecells.HtmlHiddenColDisplayType) Option {
	return func(c *Config) {
		c.hiddenColDisplayType = &value
	}
}
func WithHiddenRowDisplayType(value asposecells.HtmlHiddenRowDisplayType) Option {
	return func(c *Config) {
		c.hiddenRowDisplayType = &value
	}
}
func WithEncoding(value asposecells.EncodingType) Option {
	return func(c *Config) {
		c.encoding = &value
	}
}
func WithSaveAsSingleFile(value bool) Option {
	return func(c *Config) {
		c.saveAsSingleFile = &value
	}
}

func WithShowAllSheets(value bool) Option {
	return func(c *Config) {
		c.showAllSheets = &value
	}
}

func WithExportPageHeaders(value bool) Option {
	return func(c *Config) {
		c.exportPageHeaders = &value
	}
}

func WithExportPageFooters(value bool) Option {
	return func(c *Config) {
		c.exportPageFooters = &value
	}
}

func WithExportHiddenWorksheet(value bool) Option {
	return func(c *Config) {
		c.exportHiddenWorksheet = &value
	}
}

func WithPresentationPreference(value bool) Option {
	return func(c *Config) {
		c.presentationPreference = &value
	}
}

func WithCellCssPrefix(value string) Option {
	return func(c *Config) {
		c.cellCssPrefix = &value
	}
}

func WithTableCssId(value string) Option {
	return func(c *Config) {
		c.tableCssId = &value
	}
}

func WithIsFullPathLink(value bool) Option {
	return func(c *Config) {
		c.isFullPathLink = &value
	}
}

func WithExportWorksheetCSSSeparately(value bool) Option {
	return func(c *Config) {
		c.exportWorksheetCSSSeparately = &value
	}
}

func WithExportSimilarBorderStyle(value bool) Option {
	return func(c *Config) {
		c.exportSimilarBorderStyle = &value
	}
}

func WithMergeEmptyTdType(value asposecells.MergeEmptyTdType) Option {
	return func(c *Config) {
		c.mergeEmptyTdType = &value
	}
}
func WithExportCellCoordinate(value bool) Option {
	return func(c *Config) {
		c.exportCellCoordinate = &value
	}
}

func WithExportExtraHeadings(value bool) Option {
	return func(c *Config) {
		c.exportExtraHeadings = &value
	}
}

func WithExportRowColumnHeadings(value bool) Option {
	return func(c *Config) {
		c.exportRowColumnHeadings = &value
	}
}

func WithExportFormula(value bool) Option {
	return func(c *Config) {
		c.exportFormula = &value
	}
}

func WithAddTooltipText(value bool) Option {
	return func(c *Config) {
		c.addTooltipText = &value
	}
}

func WithExportGridLines(value bool) Option {
	return func(c *Config) {
		c.exportGridLines = &value
	}
}

func WithExportBogusRowData(value bool) Option {
	return func(c *Config) {
		c.exportBogusRowData = &value
	}
}

func WithExcludeUnusedStyles(value bool) Option {
	return func(c *Config) {
		c.excludeUnusedStyles = &value
	}
}

func WithExportDocumentProperties(value bool) Option {
	return func(c *Config) {
		c.exportDocumentProperties = &value
	}
}

func WithExportWorksheetProperties(value bool) Option {
	return func(c *Config) {
		c.exportWorksheetProperties = &value
	}
}

func WithExportWorkbookProperties(value bool) Option {
	return func(c *Config) {
		c.exportWorkbookProperties = &value
	}
}

func WithExportFrameScriptsAndProperties(value bool) Option {
	return func(c *Config) {
		c.exportFrameScriptsAndProperties = &value
	}
}

func WithExportDataOptions(value asposecells.HtmlExportDataOptions) Option {
	return func(c *Config) {
		c.exportDataOptions = &value
	}
}
func WithLinkTargetType(value asposecells.HtmlLinkTargetType) Option {
	return func(c *Config) {
		c.linkTargetType = &value
	}
}
func WithIsIECompatible(value bool) Option {
	return func(c *Config) {
		c.isIECompatible = &value
	}
}

func WithFormatDataIgnoreColumnWidth(value bool) Option {
	return func(c *Config) {
		c.formatDataIgnoreColumnWidth = &value
	}
}

func WithCalculateFormula(value bool) Option {
	return func(c *Config) {
		c.calculateFormula = &value
	}
}

func WithIsJsBrowserCompatible(value bool) Option {
	return func(c *Config) {
		c.isJsBrowserCompatible = &value
	}
}

func WithIsMobileCompatible(value bool) Option {
	return func(c *Config) {
		c.isMobileCompatible = &value
	}
}

func WithCssStyles(value string) Option {
	return func(c *Config) {
		c.cssStyles = &value
	}
}

func WithHideOverflowWrappedText(value bool) Option {
	return func(c *Config) {
		c.hideOverflowWrappedText = &value
	}
}

func WithIsBorderCollapsed(value bool) Option {
	return func(c *Config) {
		c.isBorderCollapsed = &value
	}
}

func WithEncodeEntityAsCode(value bool) Option {
	return func(c *Config) {
		c.encodeEntityAsCode = &value
	}
}

func WithOfficeMathOutputMode(value asposecells.HtmlOfficeMathOutputType) Option {
	return func(c *Config) {
		c.officeMathOutputMode = &value
	}
}
func WithCellNameAttribute(value string) Option {
	return func(c *Config) {
		c.cellNameAttribute = &value
	}
}

func WithDisableCss(value bool) Option {
	return func(c *Config) {
		c.disableCss = &value
	}
}

func WithEnableCssCustomProperties(value bool) Option {
	return func(c *Config) {
		c.enableCssCustomProperties = &value
	}
}

func WithHtmlVersion(value asposecells.HtmlVersion) Option {
	return func(c *Config) {
		c.htmlVersion = &value
	}
}
func WithSheetSet(value *asposecells.SheetSet) Option {
	return func(c *Config) {
		c.sheetSet = value
	}
}
func WithLayoutMode(value asposecells.HtmlLayoutMode) Option {
	return func(c *Config) {
		c.layoutMode = &value
	}
}
func WithEmbeddedFontType(value asposecells.HtmlEmbeddedFontType) Option {
	return func(c *Config) {
		c.embeddedFontType = &value
	}
}
func WithExportNamedRangeAnchors(value bool) Option {
	return func(c *Config) {
		c.exportNamedRangeAnchors = &value
	}
}

func WithDataBarRenderMode(value asposecells.DataBarRenderMode) Option {
	return func(c *Config) {
		c.dataBarRenderMode = &value
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
