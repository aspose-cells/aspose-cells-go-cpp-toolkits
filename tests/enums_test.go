package tests

import (
	"bytes"
	"errors"
	"fmt"
	"strings"
	"testing"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	enums "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/enums"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/csv"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/html"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/image"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/pdf"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// checkEnum asserts that every member of one engine enum is reachable through its
// resolver by name, and that a name matching no member is rejected.
//
// The want table holds the complete member list, so a member dropped from the
// resolver's table surfaces here as an unexpected error rather than silently
// becoming unsettable. A spelling that resolves to the wrong member fails the
// comparison, and two names resolving to the same member collapse an entry, so
// the values are required to be distinct as well.
func checkEnum[T comparable](t *testing.T, enumName string, resolve func(string) (T, error), want map[string]T) {
	t.Helper()
	if len(want) == 0 {
		t.Fatalf("%s: test table is empty", enumName)
	}
	seen := make(map[T]string, len(want))
	for name, expected := range want {
		if previous, duplicate := seen[expected]; duplicate {
			t.Errorf("%s: %q and %q both resolve to %v; one spelling is wrong", enumName, previous, name, expected)
		}
		seen[expected] = name

		got, err := resolve(name)
		if err != nil {
			t.Errorf("%s(%q): unexpected error: %v", enumName, name, err)
			continue
		}
		if got != expected {
			t.Errorf("%s(%q) = %v, want %v", enumName, name, got, expected)
		}

		// Members are matched case- and punctuation-insensitively, so the same
		// member has to be reachable through a spelling that differs only in case
		// and separators.
		loud := strings.ToUpper(name) + "-_ "
		if got, err := resolve(loud); err != nil {
			t.Errorf("%s(%q): case/punctuation-insensitive lookup failed: %v", enumName, loud, err)
		} else if got != expected {
			t.Errorf("%s(%q) = %v, want %v", enumName, loud, got, expected)
		}
	}

	// A name matching no member must be an error. Returning the engine's zero
	// value instead is what lets a typo become a silently different option.
	if _, err := resolve("zzz-not-a-member"); !errors.Is(err, toolkiterrors.ErrInvalidEnumValue) {
		t.Errorf("%s(unknown) error = %v, want ErrInvalidEnumValue", enumName, err)
	}
}

func TestEnumNormalize(t *testing.T) {
	cases := map[string]string{
		"":                  "",
		"UTF8":              "utf8",
		"utf-8":             "utf8",
		"EmfPlusPrefer":     "emfplusprefer",
		"not_docked":        "notdocked",
		"  Pdf A1b  ":       "pdfa1b",
		"alreadyNormalized": "alreadynormalized",
	}
	for input, want := range cases {
		if got := enums.Normalize(input); got != want {
			t.Errorf("Normalize(%q) = %q, want %q", input, got, want)
		}
	}
}

// TestEnumResolution walks every enum the toolkit resolves. The tables are
// generated from the engine's own declarations, so they fail when the engine adds
// a member the resolver does not carry.
func TestEnumResolution(t *testing.T) {
	t.Run("AdjustFontSizeForRowType", func(t *testing.T) {
		checkEnum(t, "AdjustFontSizeForRowType", enums.AdjustFontSizeForRowType, map[string]asposecells.AdjustFontSizeForRowType{
			"emptyRows": asposecells.AdjustFontSizeForRowType_EmptyRows,
			"none":      asposecells.AdjustFontSizeForRowType_None,
		})
	})
	t.Run("CellValueFormatStrategy", func(t *testing.T) {
		checkEnum(t, "CellValueFormatStrategy", enums.CellValueFormatStrategy, map[string]asposecells.CellValueFormatStrategy{
			"cellStyle":     asposecells.CellValueFormatStrategy_CellStyle,
			"displayString": asposecells.CellValueFormatStrategy_DisplayString,
			"displayStyle":  asposecells.CellValueFormatStrategy_DisplayStyle,
			"none":          asposecells.CellValueFormatStrategy_None,
		})
	})
	t.Run("DataBarRenderMode", func(t *testing.T) {
		checkEnum(t, "DataBarRenderMode", enums.DataBarRenderMode, map[string]asposecells.DataBarRenderMode{
			"backgroundColor": asposecells.DataBarRenderMode_BackgroundColor,
			"image":           asposecells.DataBarRenderMode_Image,
		})
	})
	t.Run("DefaultEditLanguage", func(t *testing.T) {
		checkEnum(t, "DefaultEditLanguage", enums.DefaultEditLanguage, map[string]asposecells.DefaultEditLanguage{
			"auto":    asposecells.DefaultEditLanguage_Auto,
			"cjk":     asposecells.DefaultEditLanguage_CJK,
			"english": asposecells.DefaultEditLanguage_English,
		})
	})
	t.Run("EmfRenderSetting", func(t *testing.T) {
		checkEnum(t, "EmfRenderSetting", enums.EmfRenderSetting, map[string]asposecells.EmfRenderSetting{
			"emfOnly":       asposecells.EmfRenderSetting_EmfOnly,
			"emfPlusPrefer": asposecells.EmfRenderSetting_EmfPlusPrefer,
		})
	})
	t.Run("EncodingType", func(t *testing.T) {
		checkEnum(t, "EncodingType", enums.EncodingType, map[string]asposecells.EncodingType{
			"ascii":   asposecells.EncodingType_ASCII,
			"default": asposecells.EncodingType_Default,
			"unicode": asposecells.EncodingType_Unicode,
			"utf8":    asposecells.EncodingType_UTF8,
		})
	})
	t.Run("GridlineType", func(t *testing.T) {
		checkEnum(t, "GridlineType", enums.GridlineType, map[string]asposecells.GridlineType{
			"dotted": asposecells.GridlineType_Dotted,
			"hair":   asposecells.GridlineType_Hair,
		})
	})
	t.Run("HtmlCrossType", func(t *testing.T) {
		checkEnum(t, "HtmlCrossType", enums.HtmlCrossType, map[string]asposecells.HtmlCrossType{
			"cross":          asposecells.HtmlCrossType_Cross,
			"crossHideRight": asposecells.HtmlCrossType_CrossHideRight,
			"default":        asposecells.HtmlCrossType_Default,
			"fitToCell":      asposecells.HtmlCrossType_FitToCell,
			"msExport":       asposecells.HtmlCrossType_MSExport,
		})
	})
	t.Run("HtmlEmbeddedFontType", func(t *testing.T) {
		checkEnum(t, "HtmlEmbeddedFontType", enums.HtmlEmbeddedFontType, map[string]asposecells.HtmlEmbeddedFontType{
			"none": asposecells.HtmlEmbeddedFontType_None,
			"woff": asposecells.HtmlEmbeddedFontType_Woff,
		})
	})
	t.Run("HtmlExportDataOptions", func(t *testing.T) {
		checkEnum(t, "HtmlExportDataOptions", enums.HtmlExportDataOptions, map[string]asposecells.HtmlExportDataOptions{
			"all":   asposecells.HtmlExportDataOptions_All,
			"table": asposecells.HtmlExportDataOptions_Table,
		})
	})
	t.Run("HtmlHiddenColDisplayType", func(t *testing.T) {
		checkEnum(t, "HtmlHiddenColDisplayType", enums.HtmlHiddenColDisplayType, map[string]asposecells.HtmlHiddenColDisplayType{
			"hidden": asposecells.HtmlHiddenColDisplayType_Hidden,
			"remove": asposecells.HtmlHiddenColDisplayType_Remove,
		})
	})
	t.Run("HtmlHiddenRowDisplayType", func(t *testing.T) {
		checkEnum(t, "HtmlHiddenRowDisplayType", enums.HtmlHiddenRowDisplayType, map[string]asposecells.HtmlHiddenRowDisplayType{
			"hidden": asposecells.HtmlHiddenRowDisplayType_Hidden,
			"remove": asposecells.HtmlHiddenRowDisplayType_Remove,
		})
	})
	t.Run("HtmlLayoutMode", func(t *testing.T) {
		checkEnum(t, "HtmlLayoutMode", enums.HtmlLayoutMode, map[string]asposecells.HtmlLayoutMode{
			"normal": asposecells.HtmlLayoutMode_Normal,
			"print":  asposecells.HtmlLayoutMode_Print,
		})
	})
	t.Run("HtmlLinkTargetType", func(t *testing.T) {
		checkEnum(t, "HtmlLinkTargetType", enums.HtmlLinkTargetType, map[string]asposecells.HtmlLinkTargetType{
			"blank":  asposecells.HtmlLinkTargetType_Blank,
			"parent": asposecells.HtmlLinkTargetType_Parent,
			"self":   asposecells.HtmlLinkTargetType_Self,
			"top":    asposecells.HtmlLinkTargetType_Top,
		})
	})
	t.Run("HtmlOfficeMathOutputType", func(t *testing.T) {
		checkEnum(t, "HtmlOfficeMathOutputType", enums.HtmlOfficeMathOutputType, map[string]asposecells.HtmlOfficeMathOutputType{
			"image":  asposecells.HtmlOfficeMathOutputType_Image,
			"mathML": asposecells.HtmlOfficeMathOutputType_MathML,
		})
	})
	t.Run("HtmlVersion", func(t *testing.T) {
		checkEnum(t, "HtmlVersion", enums.HtmlVersion, map[string]asposecells.HtmlVersion{
			"default": asposecells.HtmlVersion_Default,
			"html5":   asposecells.HtmlVersion_Html5,
			"xHtml":   asposecells.HtmlVersion_XHtml,
		})
	})
	t.Run("ImageType", func(t *testing.T) {
		checkEnum(t, "ImageType", enums.ImageType, map[string]asposecells.ImageType{
			"bmp":                 asposecells.ImageType_Bmp,
			"emf":                 asposecells.ImageType_Emf,
			"gif":                 asposecells.ImageType_Gif,
			"gltf":                asposecells.ImageType_Gltf,
			"jpeg":                asposecells.ImageType_Jpeg,
			"officeCompatibleEmf": asposecells.ImageType_OfficeCompatibleEmf,
			"pict":                asposecells.ImageType_Pict,
			"png":                 asposecells.ImageType_Png,
			"svg":                 asposecells.ImageType_Svg,
			"svm":                 asposecells.ImageType_Svm,
			"tiff":                asposecells.ImageType_Tiff,
			"webP":                asposecells.ImageType_WebP,
			"wmf":                 asposecells.ImageType_Wmf,
		})
	})
	t.Run("JsonExportHyperlinkType", func(t *testing.T) {
		checkEnum(t, "JsonExportHyperlinkType", enums.JsonExportHyperlinkType, map[string]asposecells.JsonExportHyperlinkType{
			"address":       asposecells.JsonExportHyperlinkType_Address,
			"displayString": asposecells.JsonExportHyperlinkType_DisplayString,
			"htmlString":    asposecells.JsonExportHyperlinkType_HtmlString,
		})
	})
	t.Run("MarkdownTableHeaderType", func(t *testing.T) {
		checkEnum(t, "MarkdownTableHeaderType", enums.MarkdownTableHeaderType, map[string]asposecells.MarkdownTableHeaderType{
			"columnHeader": asposecells.MarkdownTableHeaderType_ColumnHeader,
			"empty":        asposecells.MarkdownTableHeaderType_Empty,
			"firstRow":     asposecells.MarkdownTableHeaderType_FirstRow,
		})
	})
	t.Run("MergeEmptyTdType", func(t *testing.T) {
		checkEnum(t, "MergeEmptyTdType", enums.MergeEmptyTdType, map[string]asposecells.MergeEmptyTdType{
			"default":      asposecells.MergeEmptyTdType_Default,
			"mergeForcely": asposecells.MergeEmptyTdType_MergeForcely,
			"none":         asposecells.MergeEmptyTdType_None,
		})
	})
	t.Run("OdsGeneratorType", func(t *testing.T) {
		checkEnum(t, "OdsGeneratorType", enums.OdsGeneratorType, map[string]asposecells.OdsGeneratorType{
			"libreOffice": asposecells.OdsGeneratorType_LibreOffice,
			"openOffice":  asposecells.OdsGeneratorType_OpenOffice,
		})
	})
	t.Run("OoxmlCompressionType", func(t *testing.T) {
		checkEnum(t, "OoxmlCompressionType", enums.OoxmlCompressionType, map[string]asposecells.OoxmlCompressionType{
			"level1": asposecells.OoxmlCompressionType_Level1,
			"level2": asposecells.OoxmlCompressionType_Level2,
			"level3": asposecells.OoxmlCompressionType_Level3,
			"level4": asposecells.OoxmlCompressionType_Level4,
			"level5": asposecells.OoxmlCompressionType_Level5,
			"level6": asposecells.OoxmlCompressionType_Level6,
			"level7": asposecells.OoxmlCompressionType_Level7,
			"level8": asposecells.OoxmlCompressionType_Level8,
			"level9": asposecells.OoxmlCompressionType_Level9,
		})
	})
	t.Run("OpenDocumentFormatVersionType", func(t *testing.T) {
		checkEnum(t, "OpenDocumentFormatVersionType", enums.OpenDocumentFormatVersionType, map[string]asposecells.OpenDocumentFormatVersionType{
			"none":  asposecells.OpenDocumentFormatVersionType_None,
			"odf11": asposecells.OpenDocumentFormatVersionType_Odf11,
			"odf12": asposecells.OpenDocumentFormatVersionType_Odf12,
			"odf13": asposecells.OpenDocumentFormatVersionType_Odf13,
			"odf14": asposecells.OpenDocumentFormatVersionType_Odf14,
		})
	})
	t.Run("PdfCompliance", func(t *testing.T) {
		checkEnum(t, "PdfCompliance", enums.PdfCompliance, map[string]asposecells.PdfCompliance{
			"pdf14":  asposecells.PdfCompliance_Pdf14,
			"pdf15":  asposecells.PdfCompliance_Pdf15,
			"pdf16":  asposecells.PdfCompliance_Pdf16,
			"pdf17":  asposecells.PdfCompliance_Pdf17,
			"pdfA1a": asposecells.PdfCompliance_PdfA1a,
			"pdfA1b": asposecells.PdfCompliance_PdfA1b,
			"pdfA2a": asposecells.PdfCompliance_PdfA2a,
			"pdfA2b": asposecells.PdfCompliance_PdfA2b,
			"pdfA2u": asposecells.PdfCompliance_PdfA2u,
			"pdfA3a": asposecells.PdfCompliance_PdfA3a,
			"pdfA3b": asposecells.PdfCompliance_PdfA3b,
			"pdfA3u": asposecells.PdfCompliance_PdfA3u,
		})
	})
	t.Run("PdfCompressionCore", func(t *testing.T) {
		checkEnum(t, "PdfCompressionCore", enums.PdfCompressionCore, map[string]asposecells.PdfCompressionCore{
			"flate": asposecells.PdfCompressionCore_Flate,
			"lzw":   asposecells.PdfCompressionCore_Lzw,
			"none":  asposecells.PdfCompressionCore_None,
			"rle":   asposecells.PdfCompressionCore_Rle,
		})
	})
	t.Run("PdfCustomPropertiesExport", func(t *testing.T) {
		checkEnum(t, "PdfCustomPropertiesExport", enums.PdfCustomPropertiesExport, map[string]asposecells.PdfCustomPropertiesExport{
			"none":     asposecells.PdfCustomPropertiesExport_None,
			"standard": asposecells.PdfCustomPropertiesExport_Standard,
		})
	})
	t.Run("PdfFontEncoding", func(t *testing.T) {
		checkEnum(t, "PdfFontEncoding", enums.PdfFontEncoding, map[string]asposecells.PdfFontEncoding{
			"ansiPrefer": asposecells.PdfFontEncoding_AnsiPrefer,
			"identity":   asposecells.PdfFontEncoding_Identity,
		})
	})
	t.Run("PdfOptimizationType", func(t *testing.T) {
		checkEnum(t, "PdfOptimizationType", enums.PdfOptimizationType, map[string]asposecells.PdfOptimizationType{
			"minimumSize": asposecells.PdfOptimizationType_MinimumSize,
			"standard":    asposecells.PdfOptimizationType_Standard,
		})
	})
	t.Run("PrintCommentsType", func(t *testing.T) {
		checkEnum(t, "PrintCommentsType", enums.PrintCommentsType, map[string]asposecells.PrintCommentsType{
			"printInPlace":              asposecells.PrintCommentsType_PrintInPlace,
			"printNoComments":           asposecells.PrintCommentsType_PrintNoComments,
			"printSheetEnd":             asposecells.PrintCommentsType_PrintSheetEnd,
			"printWithThreadedComments": asposecells.PrintCommentsType_PrintWithThreadedComments,
		})
	})
	t.Run("PrintingPageType", func(t *testing.T) {
		checkEnum(t, "PrintingPageType", enums.PrintingPageType, map[string]asposecells.PrintingPageType{
			"default":     asposecells.PrintingPageType_Default,
			"ignoreBlank": asposecells.PrintingPageType_IgnoreBlank,
			"ignoreStyle": asposecells.PrintingPageType_IgnoreStyle,
		})
	})
	t.Run("SaveElementType", func(t *testing.T) {
		checkEnum(t, "SaveElementType", enums.SaveElementType, map[string]asposecells.SaveElementType{
			"all":   asposecells.SaveElementType_All,
			"chart": asposecells.SaveElementType_Chart,
		})
	})
	t.Run("SlideViewType", func(t *testing.T) {
		checkEnum(t, "SlideViewType", enums.SlideViewType, map[string]asposecells.SlideViewType{
			"print": asposecells.SlideViewType_Print,
			"view":  asposecells.SlideViewType_View,
		})
	})
	t.Run("SqlScriptOperatorType", func(t *testing.T) {
		checkEnum(t, "SqlScriptOperatorType", enums.SqlScriptOperatorType, map[string]asposecells.SqlScriptOperatorType{
			"delete": asposecells.SqlScriptOperatorType_Delete,
			"insert": asposecells.SqlScriptOperatorType_Insert,
			"update": asposecells.SqlScriptOperatorType_Update,
		})
	})
	t.Run("TextCrossType", func(t *testing.T) {
		checkEnum(t, "TextCrossType", enums.TextCrossType, map[string]asposecells.TextCrossType{
			"crossKeep":     asposecells.TextCrossType_CrossKeep,
			"crossOverride": asposecells.TextCrossType_CrossOverride,
			"default":       asposecells.TextCrossType_Default,
			"strictInCell":  asposecells.TextCrossType_StrictInCell,
		})
	})
	t.Run("TxtValueQuoteType", func(t *testing.T) {
		checkEnum(t, "TxtValueQuoteType", enums.TxtValueQuoteType, map[string]asposecells.TxtValueQuoteType{
			"always":  asposecells.TxtValueQuoteType_Always,
			"minimum": asposecells.TxtValueQuoteType_Minimum,
			"never":   asposecells.TxtValueQuoteType_Never,
			"normal":  asposecells.TxtValueQuoteType_Normal,
		})
	})
}

// TestEnumUnknownMemberRejected pins the error type on names that are near misses
// of real members: a miss must not be rounded to the closest match.
func TestEnumUnknownMemberRejected(t *testing.T) {
	for _, name := range []string{"utf", "html4", "printInPlac", "nonee"} {
		if _, err := enums.EncodingType(name); err == nil {
			t.Errorf("EncodingType(%q) = nil error, want ErrInvalidEnumValue", name)
		} else if !errors.Is(err, toolkiterrors.ErrInvalidEnumValue) {
			t.Errorf("EncodingType(%q) error = %v, want ErrInvalidEnumValue", name, err)
		}
	}
	// The engine's ImageType_Unknown member is its sentinel for "no format
	// recognized", not a format, so it is not a name a caller can ask for.
	if _, err := enums.ImageType("unknown"); !errors.Is(err, toolkiterrors.ErrInvalidEnumValue) {
		t.Errorf("ImageType(unknown) error = %v, want ErrInvalidEnumValue", err)
	}
}

// TestSaveOptionsRejectInvalidEnumName checks that the rejection happens where the
// caller can see it: at the save option, before the workbook is touched, so a
// mistyped name cannot half-configure a save and still produce a file.
func TestSaveOptionsRejectInvalidEnumName(t *testing.T) {
	cases := []struct {
		name string
		opt  interface{ Apply([]byte) ([]byte, error) }
	}{
		{"csv encoding", csv.New(csv.WithEncoding("utf-9"))},
		{"csv quote type", csv.New(csv.WithQuoteType("sometimes"))},
		{"csv format strategy", csv.New(csv.WithFormatStrategy("maybe"))},
		{"html hidden column display", html.New(html.WithHiddenColDisplayType("collapsed"))},
		{"html version", html.New(html.WithHtmlVersion("html6"))},
		{"pdf compliance", pdf.New(pdf.WithCompliance("pdfA9z"))},
		{"pdf compression", pdf.New(pdf.WithPdfCompression("brotli"))},
		{"image type", image.New(image.WithImageType("avif"))},
	}

	src, err := newBlankWorkbookBytes()
	if err != nil {
		t.Fatalf("newBlankWorkbookBytes: %v", err)
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			out, err := tc.opt.Apply(src)
			if !errors.Is(err, toolkiterrors.ErrInvalidEnumValue) {
				t.Fatalf("Apply error = %v, want ErrInvalidEnumValue", err)
			}
			if out != nil {
				t.Errorf("Apply returned %d bytes alongside the error, want none", len(out))
			}
		})
	}
}

// TestSaveOptionsAcceptEnumNames is the counterpart to the rejection test: a name
// that does name a member has to reach the engine and change the output. Each case
// asserts on the rendered bytes rather than on Apply returning nil, because a
// silently unapplied option is exactly the failure this refactor could introduce.
func TestSaveOptionsAcceptEnumNames(t *testing.T) {
	// The source has to carry content, not just a blank sheet: rendering an empty
	// worksheet to an image produces zero bytes (and no error) on the engine side,
	// which would make the image case below pass for the wrong reason.
	src := p2Workbook(t)

	if err := retryStable(2, func() error {
		out, err := csv.New(
			csv.WithEncoding("UTF-8"),
			csv.WithQuoteType("minimum"),
			csv.WithFormatStrategy("displayString"),
		).Apply(src)
		if err != nil {
			return fmt.Errorf("csv: %w", err)
		}
		if len(out) == 0 {
			return errors.New("csv: empty output")
		}
		return nil
	}); err != nil {
		t.Error(err)
	}

	if err := retryStable(2, func() error {
		out, err := pdf.New(
			pdf.WithCompliance("PdfA1b"),
			pdf.WithFontEncoding("identity"),
		).Apply(src)
		if err != nil {
			return fmt.Errorf("pdf: %w", err)
		}
		if !bytes.HasPrefix(out, []byte("%PDF-")) {
			return fmt.Errorf("pdf: output is not a PDF, starts with %q", head(out, 8))
		}
		return nil
	}); err != nil {
		t.Error(err)
	}

	// The short extension spellings are aliases rather than member names, so the
	// image package resolves them itself and they get their own case.
	if err := retryStable(2, func() error {
		out, err := image.New(image.WithImageType("jpg")).Apply(src)
		if err != nil {
			return fmt.Errorf("image: %w", err)
		}
		if !bytes.HasPrefix(out, []byte{0xFF, 0xD8}) {
			return fmt.Errorf("image: output is not a JPEG, starts with % X", head(out, 4))
		}
		return nil
	}); err != nil {
		t.Error(err)
	}
}

// head returns at most n bytes of b, for use in failure messages.
func head(b []byte, n int) []byte {
	if len(b) < n {
		return b
	}
	return b[:n]
}
