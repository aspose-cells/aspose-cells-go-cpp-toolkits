// Package html provides a SaveOption and configuration options for exporting
// spreadsheets as HTML documents.
package html

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	engine "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/engine"
	enums "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/enums"
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
	exportCommentsType               *string
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
	htmlCrossStringType              *string
	hiddenColDisplayType             *string
	hiddenRowDisplayType             *string
	encoding                         *string
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
	mergeEmptyTdType                 *string
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
	exportDataOptions                *string
	linkTargetType                   *string
	isIECompatible                   *bool
	formatDataIgnoreColumnWidth      *bool
	calculateFormula                 *bool
	isJsBrowserCompatible            *bool
	isMobileCompatible               *bool
	cssStyles                        *string
	hideOverflowWrappedText          *bool
	isBorderCollapsed                *bool
	encodeEntityAsCode               *bool
	officeMathOutputMode             *string
	cellNameAttribute                *string
	disableCss                       *bool
	enableCssCustomProperties        *bool
	htmlVersion                      *string
	sheetSet                         *asposecells.SheetSet
	layoutMode                       *string
	embeddedFontType                 *string
	exportNamedRangeAnchors          *bool
	dataBarRenderMode                *string
	saveoptions.CommonConfig
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
	engine.LockEngine()
	defer engine.UnlockEngine()
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
		value, err := enums.PrintCommentsType(*c.exportCommentsType)
		if err != nil {
			return nil, err
		}
		if err := opts.SetExportCommentsType(value); err != nil {
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
		value, err := enums.HtmlCrossType(*c.htmlCrossStringType)
		if err != nil {
			return nil, err
		}
		if err := opts.SetHtmlCrossStringType(value); err != nil {
			return nil, err
		}
	}
	if c.hiddenColDisplayType != nil {
		value, err := enums.HtmlHiddenColDisplayType(*c.hiddenColDisplayType)
		if err != nil {
			return nil, err
		}
		if err := opts.SetHiddenColDisplayType(value); err != nil {
			return nil, err
		}
	}
	if c.hiddenRowDisplayType != nil {
		value, err := enums.HtmlHiddenRowDisplayType(*c.hiddenRowDisplayType)
		if err != nil {
			return nil, err
		}
		if err := opts.SetHiddenRowDisplayType(value); err != nil {
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
		value, err := enums.MergeEmptyTdType(*c.mergeEmptyTdType)
		if err != nil {
			return nil, err
		}
		if err := opts.SetMergeEmptyTdType(value); err != nil {
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
		value, err := enums.HtmlExportDataOptions(*c.exportDataOptions)
		if err != nil {
			return nil, err
		}
		if err := opts.SetExportDataOptions(value); err != nil {
			return nil, err
		}
	}
	if c.linkTargetType != nil {
		value, err := enums.HtmlLinkTargetType(*c.linkTargetType)
		if err != nil {
			return nil, err
		}
		if err := opts.SetLinkTargetType(value); err != nil {
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
		value, err := enums.HtmlOfficeMathOutputType(*c.officeMathOutputMode)
		if err != nil {
			return nil, err
		}
		if err := opts.SetOfficeMathOutputMode(value); err != nil {
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
		value, err := enums.HtmlVersion(*c.htmlVersion)
		if err != nil {
			return nil, err
		}
		if err := opts.SetHtmlVersion(value); err != nil {
			return nil, err
		}
	}
	if c.sheetSet != nil {
		if err := opts.SetSheetSet(c.sheetSet); err != nil {
			return nil, err
		}
	}
	if c.layoutMode != nil {
		value, err := enums.HtmlLayoutMode(*c.layoutMode)
		if err != nil {
			return nil, err
		}
		if err := opts.SetLayoutMode(value); err != nil {
			return nil, err
		}
	}
	if c.embeddedFontType != nil {
		value, err := enums.HtmlEmbeddedFontType(*c.embeddedFontType)
		if err != nil {
			return nil, err
		}
		if err := opts.SetEmbeddedFontType(value); err != nil {
			return nil, err
		}
	}
	if c.exportNamedRangeAnchors != nil {
		if err := opts.SetExportNamedRangeAnchors(*c.exportNamedRangeAnchors); err != nil {
			return nil, err
		}
	}
	if c.dataBarRenderMode != nil {
		value, err := enums.DataBarRenderMode(*c.dataBarRenderMode)
		if err != nil {
			return nil, err
		}
		if err := opts.SetDataBarRenderMode(value); err != nil {
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

// WithExportCommentsType sets the export comments type: "printInPlace",
// "printNoComments", "printSheetEnd", or "printWithThreadedComments".
func WithExportCommentsType(value string) Option {
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

func WithParseHtmlTagInCell(value bool) Option {
	return func(c *Config) {
		c.parseHtmlTagInCell = &value
	}
}

// WithHtmlCrossStringType sets the HTML cross string type: "cross",
// "crossHideRight", "default", "fitToCell", or "msExport".
func WithHtmlCrossStringType(value string) Option {
	return func(c *Config) {
		c.htmlCrossStringType = &value
	}
}

// WithHiddenColDisplayType sets the hidden col display type: "hidden" or
// "remove".
func WithHiddenColDisplayType(value string) Option {
	return func(c *Config) {
		c.hiddenColDisplayType = &value
	}
}

// WithHiddenRowDisplayType sets the hidden row display type: "hidden" or
// "remove".
func WithHiddenRowDisplayType(value string) Option {
	return func(c *Config) {
		c.hiddenRowDisplayType = &value
	}
}

// WithEncoding sets the encoding: "ascii", "default", "unicode", or "utf8".
func WithEncoding(value string) Option {
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

// WithMergeEmptyTdType sets the merge empty TD type: "default",
// "mergeForcely", or "none".
func WithMergeEmptyTdType(value string) Option {
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

// WithExportDataOptions sets the export data options: "all" or "table".
func WithExportDataOptions(value string) Option {
	return func(c *Config) {
		c.exportDataOptions = &value
	}
}

// WithLinkTargetType sets the link target type: "blank", "parent", "self",
// or "top".
func WithLinkTargetType(value string) Option {
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

// WithOfficeMathOutputMode sets the office math output mode: "image" or
// "mathML".
func WithOfficeMathOutputMode(value string) Option {
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

// WithHtmlVersion sets the HTML version: "default", "html5", or "xHtml".
func WithHtmlVersion(value string) Option {
	return func(c *Config) {
		c.htmlVersion = &value
	}
}

// Disabled: this option names an engine type, which the toolkit's public API
// must not do — a caller who set it would be tied to the binding
// (docs/design.md §11). Restore it by taking the toolkit-native value instead,
// the way json.WithExportArea takes an "A1:C3" string. The original
// declaration follows verbatim.
/*
func WithSheetSet(value *asposecells.SheetSet) Option {
	return func(c *Config) {
		c.sheetSet = value
	}
}
*/

// WithLayoutMode sets the layout mode: "normal" or "print".
func WithLayoutMode(value string) Option {
	return func(c *Config) {
		c.layoutMode = &value
	}
}

// WithEmbeddedFontType sets the embedded font type: "none" or "woff".
func WithEmbeddedFontType(value string) Option {
	return func(c *Config) {
		c.embeddedFontType = &value
	}
}
func WithExportNamedRangeAnchors(value bool) Option {
	return func(c *Config) {
		c.exportNamedRangeAnchors = &value
	}
}

// WithDataBarRenderMode sets the data bar render mode: "backgroundColor" or
// "image".
func WithDataBarRenderMode(value string) Option {
	return func(c *Config) {
		c.dataBarRenderMode = &value
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
