// Package enums maps toolkit-native string names to the binding's enum values,
// so that no public save option has to name an asposecells.* type.
//
// The binding's enums are int32 types with members named like EncodingType_UTF8.
// Exposing them through With* options would tie callers to the binding and make
// an engine upgrade a breaking change for them, so save options accept a string
// and resolve it here instead.
//
// Names are matched the way the rest of the toolkit matches them: case- and
// punctuation-insensitively, so "utf8", "UTF8" and "utf-8" are one name. The
// canonical spelling of each member is its constant suffix in lower camel case
// (UTF8 -> "utf8", EmfPlusPrefer -> "emfPlusPrefer"), and that spelling is what
// the option documentation quotes. The tables below are keyed by the normalized
// form, because normalized is what a lookup has to match.
//
// An unrecognized name is an error, never a silent default. The engine
// enumerates its own members, so a caller who misspells one has asked for
// something that does not exist, and quietly substituting another value is how
// a mistyped option becomes a wrong document.
package enums

import (
	"fmt"
	"strings"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// Normalize lowercases name and drops every character that is not a letter or a
// digit, reducing the calling convention above to one comparison. It is the same
// rule editor applies to its own string enums.
func Normalize(name string) string {
	var b strings.Builder
	b.Grow(len(name))
	for _, r := range name {
		switch {
		case r >= 'a' && r <= 'z', r >= '0' && r <= '9':
			b.WriteRune(r)
		case r >= 'A' && r <= 'Z':
			b.WriteRune(r + ('a' - 'A'))
		}
	}
	return b.String()
}

// adjustFontSizeForRowTypeNames maps a normalized name to the engine's AdjustFontSizeForRowType.
var adjustFontSizeForRowTypeNames = map[string]asposecells.AdjustFontSizeForRowType{
	"emptyrows": asposecells.AdjustFontSizeForRowType_EmptyRows,
	"none":      asposecells.AdjustFontSizeForRowType_None,
}

// AdjustFontSizeForRowType resolves an AdjustFontSizeForRowType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func AdjustFontSizeForRowType(name string) (asposecells.AdjustFontSizeForRowType, error) {
	if v, ok := adjustFontSizeForRowTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown adjustFontSizeForRowType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// cellValueFormatStrategyNames maps a normalized name to the engine's CellValueFormatStrategy.
var cellValueFormatStrategyNames = map[string]asposecells.CellValueFormatStrategy{
	"cellstyle":     asposecells.CellValueFormatStrategy_CellStyle,
	"displaystring": asposecells.CellValueFormatStrategy_DisplayString,
	"displaystyle":  asposecells.CellValueFormatStrategy_DisplayStyle,
	"none":          asposecells.CellValueFormatStrategy_None,
}

// CellValueFormatStrategy resolves a CellValueFormatStrategy name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func CellValueFormatStrategy(name string) (asposecells.CellValueFormatStrategy, error) {
	if v, ok := cellValueFormatStrategyNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown cellValueFormatStrategy %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// dataBarRenderModeNames maps a normalized name to the engine's DataBarRenderMode.
var dataBarRenderModeNames = map[string]asposecells.DataBarRenderMode{
	"backgroundcolor": asposecells.DataBarRenderMode_BackgroundColor,
	"image":           asposecells.DataBarRenderMode_Image,
}

// DataBarRenderMode resolves a DataBarRenderMode name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func DataBarRenderMode(name string) (asposecells.DataBarRenderMode, error) {
	if v, ok := dataBarRenderModeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown dataBarRenderMode %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// defaultEditLanguageNames maps a normalized name to the engine's DefaultEditLanguage.
var defaultEditLanguageNames = map[string]asposecells.DefaultEditLanguage{
	"auto":    asposecells.DefaultEditLanguage_Auto,
	"cjk":     asposecells.DefaultEditLanguage_CJK,
	"english": asposecells.DefaultEditLanguage_English,
}

// DefaultEditLanguage resolves a DefaultEditLanguage name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func DefaultEditLanguage(name string) (asposecells.DefaultEditLanguage, error) {
	if v, ok := defaultEditLanguageNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown defaultEditLanguage %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// emfRenderSettingNames maps a normalized name to the engine's EmfRenderSetting.
var emfRenderSettingNames = map[string]asposecells.EmfRenderSetting{
	"emfonly":       asposecells.EmfRenderSetting_EmfOnly,
	"emfplusprefer": asposecells.EmfRenderSetting_EmfPlusPrefer,
}

// EmfRenderSetting resolves an EmfRenderSetting name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func EmfRenderSetting(name string) (asposecells.EmfRenderSetting, error) {
	if v, ok := emfRenderSettingNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown emfRenderSetting %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// encodingTypeNames maps a normalized name to the engine's EncodingType.
var encodingTypeNames = map[string]asposecells.EncodingType{
	"ascii":   asposecells.EncodingType_ASCII,
	"default": asposecells.EncodingType_Default,
	"unicode": asposecells.EncodingType_Unicode,
	"utf8":    asposecells.EncodingType_UTF8,
}

// EncodingType resolves an EncodingType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func EncodingType(name string) (asposecells.EncodingType, error) {
	if v, ok := encodingTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown encodingType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// gridlineTypeNames maps a normalized name to the engine's GridlineType.
var gridlineTypeNames = map[string]asposecells.GridlineType{
	"dotted": asposecells.GridlineType_Dotted,
	"hair":   asposecells.GridlineType_Hair,
}

// GridlineType resolves a GridlineType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func GridlineType(name string) (asposecells.GridlineType, error) {
	if v, ok := gridlineTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown gridlineType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// htmlCrossTypeNames maps a normalized name to the engine's HtmlCrossType.
var htmlCrossTypeNames = map[string]asposecells.HtmlCrossType{
	"cross":          asposecells.HtmlCrossType_Cross,
	"crosshideright": asposecells.HtmlCrossType_CrossHideRight,
	"default":        asposecells.HtmlCrossType_Default,
	"fittocell":      asposecells.HtmlCrossType_FitToCell,
	"msexport":       asposecells.HtmlCrossType_MSExport,
}

// HtmlCrossType resolves a HtmlCrossType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func HtmlCrossType(name string) (asposecells.HtmlCrossType, error) {
	if v, ok := htmlCrossTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown htmlCrossType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// htmlEmbeddedFontTypeNames maps a normalized name to the engine's HtmlEmbeddedFontType.
var htmlEmbeddedFontTypeNames = map[string]asposecells.HtmlEmbeddedFontType{
	"none": asposecells.HtmlEmbeddedFontType_None,
	"woff": asposecells.HtmlEmbeddedFontType_Woff,
}

// HtmlEmbeddedFontType resolves a HtmlEmbeddedFontType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func HtmlEmbeddedFontType(name string) (asposecells.HtmlEmbeddedFontType, error) {
	if v, ok := htmlEmbeddedFontTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown htmlEmbeddedFontType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// htmlExportDataOptionsNames maps a normalized name to the engine's HtmlExportDataOptions.
var htmlExportDataOptionsNames = map[string]asposecells.HtmlExportDataOptions{
	"all":   asposecells.HtmlExportDataOptions_All,
	"table": asposecells.HtmlExportDataOptions_Table,
}

// HtmlExportDataOptions resolves a HtmlExportDataOptions name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func HtmlExportDataOptions(name string) (asposecells.HtmlExportDataOptions, error) {
	if v, ok := htmlExportDataOptionsNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown htmlExportDataOptions %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// htmlHiddenColDisplayTypeNames maps a normalized name to the engine's HtmlHiddenColDisplayType.
var htmlHiddenColDisplayTypeNames = map[string]asposecells.HtmlHiddenColDisplayType{
	"hidden": asposecells.HtmlHiddenColDisplayType_Hidden,
	"remove": asposecells.HtmlHiddenColDisplayType_Remove,
}

// HtmlHiddenColDisplayType resolves a HtmlHiddenColDisplayType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func HtmlHiddenColDisplayType(name string) (asposecells.HtmlHiddenColDisplayType, error) {
	if v, ok := htmlHiddenColDisplayTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown htmlHiddenColDisplayType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// htmlHiddenRowDisplayTypeNames maps a normalized name to the engine's HtmlHiddenRowDisplayType.
var htmlHiddenRowDisplayTypeNames = map[string]asposecells.HtmlHiddenRowDisplayType{
	"hidden": asposecells.HtmlHiddenRowDisplayType_Hidden,
	"remove": asposecells.HtmlHiddenRowDisplayType_Remove,
}

// HtmlHiddenRowDisplayType resolves a HtmlHiddenRowDisplayType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func HtmlHiddenRowDisplayType(name string) (asposecells.HtmlHiddenRowDisplayType, error) {
	if v, ok := htmlHiddenRowDisplayTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown htmlHiddenRowDisplayType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// htmlLayoutModeNames maps a normalized name to the engine's HtmlLayoutMode.
var htmlLayoutModeNames = map[string]asposecells.HtmlLayoutMode{
	"normal": asposecells.HtmlLayoutMode_Normal,
	"print":  asposecells.HtmlLayoutMode_Print,
}

// HtmlLayoutMode resolves a HtmlLayoutMode name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func HtmlLayoutMode(name string) (asposecells.HtmlLayoutMode, error) {
	if v, ok := htmlLayoutModeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown htmlLayoutMode %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// htmlLinkTargetTypeNames maps a normalized name to the engine's HtmlLinkTargetType.
var htmlLinkTargetTypeNames = map[string]asposecells.HtmlLinkTargetType{
	"blank":  asposecells.HtmlLinkTargetType_Blank,
	"parent": asposecells.HtmlLinkTargetType_Parent,
	"self":   asposecells.HtmlLinkTargetType_Self,
	"top":    asposecells.HtmlLinkTargetType_Top,
}

// HtmlLinkTargetType resolves a HtmlLinkTargetType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func HtmlLinkTargetType(name string) (asposecells.HtmlLinkTargetType, error) {
	if v, ok := htmlLinkTargetTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown htmlLinkTargetType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// htmlOfficeMathOutputTypeNames maps a normalized name to the engine's HtmlOfficeMathOutputType.
var htmlOfficeMathOutputTypeNames = map[string]asposecells.HtmlOfficeMathOutputType{
	"image":  asposecells.HtmlOfficeMathOutputType_Image,
	"mathml": asposecells.HtmlOfficeMathOutputType_MathML,
}

// HtmlOfficeMathOutputType resolves a HtmlOfficeMathOutputType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func HtmlOfficeMathOutputType(name string) (asposecells.HtmlOfficeMathOutputType, error) {
	if v, ok := htmlOfficeMathOutputTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown htmlOfficeMathOutputType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// htmlVersionNames maps a normalized name to the engine's HtmlVersion.
var htmlVersionNames = map[string]asposecells.HtmlVersion{
	"default": asposecells.HtmlVersion_Default,
	"html5":   asposecells.HtmlVersion_Html5,
	"xhtml":   asposecells.HtmlVersion_XHtml,
}

// HtmlVersion resolves a HtmlVersion name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func HtmlVersion(name string) (asposecells.HtmlVersion, error) {
	if v, ok := htmlVersionNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown htmlVersion %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// imageTypeNames maps a normalized name to the engine's ImageType.
// ImageType_Unknown is deliberately absent. It is the engine's sentinel for
// "no format was recognized", not a format, so there is nothing a caller could
// mean by asking for it — and leaving it out is what lets a misspelled format be
// reported instead of silently producing an unrecognized file.
var imageTypeNames = map[string]asposecells.ImageType{
	"bmp":                 asposecells.ImageType_Bmp,
	"emf":                 asposecells.ImageType_Emf,
	"gif":                 asposecells.ImageType_Gif,
	"gltf":                asposecells.ImageType_Gltf,
	"jpeg":                asposecells.ImageType_Jpeg,
	"officecompatibleemf": asposecells.ImageType_OfficeCompatibleEmf,
	"pict":                asposecells.ImageType_Pict,
	"png":                 asposecells.ImageType_Png,
	"svg":                 asposecells.ImageType_Svg,
	"svm":                 asposecells.ImageType_Svm,
	"tiff":                asposecells.ImageType_Tiff,
	"webp":                asposecells.ImageType_WebP,
	"wmf":                 asposecells.ImageType_Wmf,
}

// ImageType resolves an ImageType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func ImageType(name string) (asposecells.ImageType, error) {
	if v, ok := imageTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown imageType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// jsonExportHyperlinkTypeNames maps a normalized name to the engine's JsonExportHyperlinkType.
var jsonExportHyperlinkTypeNames = map[string]asposecells.JsonExportHyperlinkType{
	"address":       asposecells.JsonExportHyperlinkType_Address,
	"displaystring": asposecells.JsonExportHyperlinkType_DisplayString,
	"htmlstring":    asposecells.JsonExportHyperlinkType_HtmlString,
}

// JsonExportHyperlinkType resolves a JsonExportHyperlinkType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func JsonExportHyperlinkType(name string) (asposecells.JsonExportHyperlinkType, error) {
	if v, ok := jsonExportHyperlinkTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown jsonExportHyperlinkType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// markdownTableHeaderTypeNames maps a normalized name to the engine's MarkdownTableHeaderType.
var markdownTableHeaderTypeNames = map[string]asposecells.MarkdownTableHeaderType{
	"columnheader": asposecells.MarkdownTableHeaderType_ColumnHeader,
	"empty":        asposecells.MarkdownTableHeaderType_Empty,
	"firstrow":     asposecells.MarkdownTableHeaderType_FirstRow,
}

// MarkdownTableHeaderType resolves a MarkdownTableHeaderType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func MarkdownTableHeaderType(name string) (asposecells.MarkdownTableHeaderType, error) {
	if v, ok := markdownTableHeaderTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown markdownTableHeaderType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// mergeEmptyTdTypeNames maps a normalized name to the engine's MergeEmptyTdType.
var mergeEmptyTdTypeNames = map[string]asposecells.MergeEmptyTdType{
	"default":      asposecells.MergeEmptyTdType_Default,
	"mergeforcely": asposecells.MergeEmptyTdType_MergeForcely,
	"none":         asposecells.MergeEmptyTdType_None,
}

// MergeEmptyTdType resolves a MergeEmptyTdType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func MergeEmptyTdType(name string) (asposecells.MergeEmptyTdType, error) {
	if v, ok := mergeEmptyTdTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown mergeEmptyTdType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// odsGeneratorTypeNames maps a normalized name to the engine's OdsGeneratorType.
var odsGeneratorTypeNames = map[string]asposecells.OdsGeneratorType{
	"libreoffice": asposecells.OdsGeneratorType_LibreOffice,
	"openoffice":  asposecells.OdsGeneratorType_OpenOffice,
}

// OdsGeneratorType resolves an OdsGeneratorType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func OdsGeneratorType(name string) (asposecells.OdsGeneratorType, error) {
	if v, ok := odsGeneratorTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown odsGeneratorType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// ooxmlCompressionTypeNames maps a normalized name to the engine's OoxmlCompressionType.
var ooxmlCompressionTypeNames = map[string]asposecells.OoxmlCompressionType{
	"level1": asposecells.OoxmlCompressionType_Level1,
	"level2": asposecells.OoxmlCompressionType_Level2,
	"level3": asposecells.OoxmlCompressionType_Level3,
	"level4": asposecells.OoxmlCompressionType_Level4,
	"level5": asposecells.OoxmlCompressionType_Level5,
	"level6": asposecells.OoxmlCompressionType_Level6,
	"level7": asposecells.OoxmlCompressionType_Level7,
	"level8": asposecells.OoxmlCompressionType_Level8,
	"level9": asposecells.OoxmlCompressionType_Level9,
}

// OoxmlCompressionType resolves an OoxmlCompressionType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func OoxmlCompressionType(name string) (asposecells.OoxmlCompressionType, error) {
	if v, ok := ooxmlCompressionTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown ooxmlCompressionType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// openDocumentFormatVersionTypeNames maps a normalized name to the engine's OpenDocumentFormatVersionType.
var openDocumentFormatVersionTypeNames = map[string]asposecells.OpenDocumentFormatVersionType{
	"none":  asposecells.OpenDocumentFormatVersionType_None,
	"odf11": asposecells.OpenDocumentFormatVersionType_Odf11,
	"odf12": asposecells.OpenDocumentFormatVersionType_Odf12,
	"odf13": asposecells.OpenDocumentFormatVersionType_Odf13,
	"odf14": asposecells.OpenDocumentFormatVersionType_Odf14,
}

// OpenDocumentFormatVersionType resolves an OpenDocumentFormatVersionType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func OpenDocumentFormatVersionType(name string) (asposecells.OpenDocumentFormatVersionType, error) {
	if v, ok := openDocumentFormatVersionTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown openDocumentFormatVersionType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// pdfComplianceNames maps a normalized name to the engine's PdfCompliance.
var pdfComplianceNames = map[string]asposecells.PdfCompliance{
	"pdf14":  asposecells.PdfCompliance_Pdf14,
	"pdf15":  asposecells.PdfCompliance_Pdf15,
	"pdf16":  asposecells.PdfCompliance_Pdf16,
	"pdf17":  asposecells.PdfCompliance_Pdf17,
	"pdfa1a": asposecells.PdfCompliance_PdfA1a,
	"pdfa1b": asposecells.PdfCompliance_PdfA1b,
	"pdfa2a": asposecells.PdfCompliance_PdfA2a,
	"pdfa2b": asposecells.PdfCompliance_PdfA2b,
	"pdfa2u": asposecells.PdfCompliance_PdfA2u,
	"pdfa3a": asposecells.PdfCompliance_PdfA3a,
	"pdfa3b": asposecells.PdfCompliance_PdfA3b,
	"pdfa3u": asposecells.PdfCompliance_PdfA3u,
}

// PdfCompliance resolves a PdfCompliance name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func PdfCompliance(name string) (asposecells.PdfCompliance, error) {
	if v, ok := pdfComplianceNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown pdfCompliance %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// pdfCompressionCoreNames maps a normalized name to the engine's PdfCompressionCore.
var pdfCompressionCoreNames = map[string]asposecells.PdfCompressionCore{
	"flate": asposecells.PdfCompressionCore_Flate,
	"lzw":   asposecells.PdfCompressionCore_Lzw,
	"none":  asposecells.PdfCompressionCore_None,
	"rle":   asposecells.PdfCompressionCore_Rle,
}

// PdfCompressionCore resolves a PdfCompressionCore name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func PdfCompressionCore(name string) (asposecells.PdfCompressionCore, error) {
	if v, ok := pdfCompressionCoreNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown pdfCompressionCore %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// pdfCustomPropertiesExportNames maps a normalized name to the engine's PdfCustomPropertiesExport.
var pdfCustomPropertiesExportNames = map[string]asposecells.PdfCustomPropertiesExport{
	"none":     asposecells.PdfCustomPropertiesExport_None,
	"standard": asposecells.PdfCustomPropertiesExport_Standard,
}

// PdfCustomPropertiesExport resolves a PdfCustomPropertiesExport name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func PdfCustomPropertiesExport(name string) (asposecells.PdfCustomPropertiesExport, error) {
	if v, ok := pdfCustomPropertiesExportNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown pdfCustomPropertiesExport %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// pdfFontEncodingNames maps a normalized name to the engine's PdfFontEncoding.
var pdfFontEncodingNames = map[string]asposecells.PdfFontEncoding{
	"ansiprefer": asposecells.PdfFontEncoding_AnsiPrefer,
	"identity":   asposecells.PdfFontEncoding_Identity,
}

// PdfFontEncoding resolves a PdfFontEncoding name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func PdfFontEncoding(name string) (asposecells.PdfFontEncoding, error) {
	if v, ok := pdfFontEncodingNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown pdfFontEncoding %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// pdfOptimizationTypeNames maps a normalized name to the engine's PdfOptimizationType.
var pdfOptimizationTypeNames = map[string]asposecells.PdfOptimizationType{
	"minimumsize": asposecells.PdfOptimizationType_MinimumSize,
	"standard":    asposecells.PdfOptimizationType_Standard,
}

// PdfOptimizationType resolves a PdfOptimizationType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func PdfOptimizationType(name string) (asposecells.PdfOptimizationType, error) {
	if v, ok := pdfOptimizationTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown pdfOptimizationType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// printCommentsTypeNames maps a normalized name to the engine's PrintCommentsType.
var printCommentsTypeNames = map[string]asposecells.PrintCommentsType{
	"printinplace":              asposecells.PrintCommentsType_PrintInPlace,
	"printnocomments":           asposecells.PrintCommentsType_PrintNoComments,
	"printsheetend":             asposecells.PrintCommentsType_PrintSheetEnd,
	"printwiththreadedcomments": asposecells.PrintCommentsType_PrintWithThreadedComments,
}

// PrintCommentsType resolves a PrintCommentsType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func PrintCommentsType(name string) (asposecells.PrintCommentsType, error) {
	if v, ok := printCommentsTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown printCommentsType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// printingPageTypeNames maps a normalized name to the engine's PrintingPageType.
var printingPageTypeNames = map[string]asposecells.PrintingPageType{
	"default":     asposecells.PrintingPageType_Default,
	"ignoreblank": asposecells.PrintingPageType_IgnoreBlank,
	"ignorestyle": asposecells.PrintingPageType_IgnoreStyle,
}

// PrintingPageType resolves a PrintingPageType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func PrintingPageType(name string) (asposecells.PrintingPageType, error) {
	if v, ok := printingPageTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown printingPageType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// saveElementTypeNames maps a normalized name to the engine's SaveElementType.
var saveElementTypeNames = map[string]asposecells.SaveElementType{
	"all":   asposecells.SaveElementType_All,
	"chart": asposecells.SaveElementType_Chart,
}

// SaveElementType resolves a SaveElementType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func SaveElementType(name string) (asposecells.SaveElementType, error) {
	if v, ok := saveElementTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown saveElementType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// slideViewTypeNames maps a normalized name to the engine's SlideViewType.
var slideViewTypeNames = map[string]asposecells.SlideViewType{
	"print": asposecells.SlideViewType_Print,
	"view":  asposecells.SlideViewType_View,
}

// SlideViewType resolves a SlideViewType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func SlideViewType(name string) (asposecells.SlideViewType, error) {
	if v, ok := slideViewTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown slideViewType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// sqlScriptOperatorTypeNames maps a normalized name to the engine's SqlScriptOperatorType.
var sqlScriptOperatorTypeNames = map[string]asposecells.SqlScriptOperatorType{
	"delete": asposecells.SqlScriptOperatorType_Delete,
	"insert": asposecells.SqlScriptOperatorType_Insert,
	"update": asposecells.SqlScriptOperatorType_Update,
}

// SqlScriptOperatorType resolves a SqlScriptOperatorType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func SqlScriptOperatorType(name string) (asposecells.SqlScriptOperatorType, error) {
	if v, ok := sqlScriptOperatorTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown sqlScriptOperatorType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// textCrossTypeNames maps a normalized name to the engine's TextCrossType.
var textCrossTypeNames = map[string]asposecells.TextCrossType{
	"crosskeep":     asposecells.TextCrossType_CrossKeep,
	"crossoverride": asposecells.TextCrossType_CrossOverride,
	"default":       asposecells.TextCrossType_Default,
	"strictincell":  asposecells.TextCrossType_StrictInCell,
}

// TextCrossType resolves a TextCrossType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func TextCrossType(name string) (asposecells.TextCrossType, error) {
	if v, ok := textCrossTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown textCrossType %q", toolkiterrors.ErrInvalidEnumValue, name)
}

// txtValueQuoteTypeNames maps a normalized name to the engine's TxtValueQuoteType.
var txtValueQuoteTypeNames = map[string]asposecells.TxtValueQuoteType{
	"always":  asposecells.TxtValueQuoteType_Always,
	"minimum": asposecells.TxtValueQuoteType_Minimum,
	"never":   asposecells.TxtValueQuoteType_Never,
	"normal":  asposecells.TxtValueQuoteType_Normal,
}

// TxtValueQuoteType resolves a TxtValueQuoteType name to the engine enum. An unknown
// name is ErrInvalidEnumValue.
func TxtValueQuoteType(name string) (asposecells.TxtValueQuoteType, error) {
	if v, ok := txtValueQuoteTypeNames[Normalize(name)]; ok {
		return v, nil
	}
	return 0, fmt.Errorf("%w: unknown txtValueQuoteType %q", toolkiterrors.ErrInvalidEnumValue, name)
}
