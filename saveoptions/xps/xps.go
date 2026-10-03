// Package xps provides a SaveOption and configuration options for exporting
// spreadsheets as XPS documents.
package xps

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	color "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/color"
	engine "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/engine"
	enums "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/enums"
	saveoptions "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// Config holds the native-typed option values for Xps save options.
//
// Pointer fields carry presence semantics: a nil pointer means the option
// was never set (so the native default is kept), while a non-nil pointer
// means the caller explicitly requested the value, including zero/false.
type Config struct {
	defaultFont                       *string
	checkWorkbookDefaultFont          *bool
	checkFontCompatibility            *bool
	isFontSubstitutionCharGranularity *bool
	onePagePerSheet                   *bool
	allColumnsInOnePagePerSheet       *bool
	ignoreError                       *bool
	outputBlankPageWhenNothingToPrint *bool
	pageIndex                         *int32
	pageCount                         *int32
	printingPageType                  *string
	gridlineType                      *string
	gridlineColor                     interface{}
	textCrossType                     *string
	defaultEditLanguage               *string
	sheetSet                          *asposecells.SheetSet
	drawObjectEventHandler            *asposecells.DrawObjectEventHandler
	emfRenderSetting                  *string
	customRenderSettings              *asposecells.CustomRenderSettings
	saveoptions.CommonConfig
}

// Apply processes the given source byte slice as a Xps file and returns the converted output.
// This method satisfies the saveoptions.SaveOption (or equivalent) interface, enabling Xps-specific export logic.
//
// Parameters:
// - source: A byte slice representing the input spreadsheet or data source. The implementation may interpret
// this as an intermediate format (e.g., XLSX or CSV bytes) and convert it into Xps format.
//
// Returns:
// - []byte: The resulting Xps file content as a byte slice.
// - error: error information.
func (c *Config) Apply(source []byte) ([]byte, error) {
	engine.LockEngine()
	defer engine.UnlockEngine()
	opts, err := asposecells.NewXpsSaveOptions()
	if err != nil {
		return nil, err
	}

	if c.defaultFont != nil {
		if err := opts.SetDefaultFont(*c.defaultFont); err != nil {
			return nil, err
		}
	}
	if c.checkWorkbookDefaultFont != nil {
		if err := opts.SetCheckWorkbookDefaultFont(*c.checkWorkbookDefaultFont); err != nil {
			return nil, err
		}
	}
	if c.checkFontCompatibility != nil {
		if err := opts.SetCheckFontCompatibility(*c.checkFontCompatibility); err != nil {
			return nil, err
		}
	}
	if c.isFontSubstitutionCharGranularity != nil {
		if err := opts.SetIsFontSubstitutionCharGranularity(*c.isFontSubstitutionCharGranularity); err != nil {
			return nil, err
		}
	}
	if c.onePagePerSheet != nil {
		if err := opts.SetOnePagePerSheet(*c.onePagePerSheet); err != nil {
			return nil, err
		}
	}
	if c.allColumnsInOnePagePerSheet != nil {
		if err := opts.SetAllColumnsInOnePagePerSheet(*c.allColumnsInOnePagePerSheet); err != nil {
			return nil, err
		}
	}
	if c.ignoreError != nil {
		if err := opts.SetIgnoreError(*c.ignoreError); err != nil {
			return nil, err
		}
	}
	if c.outputBlankPageWhenNothingToPrint != nil {
		if err := opts.SetOutputBlankPageWhenNothingToPrint(*c.outputBlankPageWhenNothingToPrint); err != nil {
			return nil, err
		}
	}
	if c.pageIndex != nil {
		if err := opts.SetPageIndex(*c.pageIndex); err != nil {
			return nil, err
		}
	}
	if c.pageCount != nil {
		if err := opts.SetPageCount(*c.pageCount); err != nil {
			return nil, err
		}
	}
	if c.printingPageType != nil {
		value, err := enums.PrintingPageType(*c.printingPageType)
		if err != nil {
			return nil, err
		}
		if err := opts.SetPrintingPageType(value); err != nil {
			return nil, err
		}
	}
	if c.gridlineType != nil {
		value, err := enums.GridlineType(*c.gridlineType)
		if err != nil {
			return nil, err
		}
		if err := opts.SetGridlineType(value); err != nil {
			return nil, err
		}
	}
	if c.gridlineColor != nil {
		gridlineColor, owned, err := color.Resolve(c.gridlineColor)
		if err != nil {
			return nil, err
		}
		if owned {
			// The engine's Color carries no finalizer, so a Color the toolkit
			// creates is the toolkit's to release; one the caller passed is not.
			// This runs after Apply returns - that is, after the save - because
			// the engine may read the pointer up until then.
			defer asposecells.DeleteColor(gridlineColor)
		}
		if err := opts.SetGridlineColor(gridlineColor); err != nil {
			return nil, err
		}
	}
	if c.textCrossType != nil {
		value, err := enums.TextCrossType(*c.textCrossType)
		if err != nil {
			return nil, err
		}
		if err := opts.SetTextCrossType(value); err != nil {
			return nil, err
		}
	}
	if c.defaultEditLanguage != nil {
		value, err := enums.DefaultEditLanguage(*c.defaultEditLanguage)
		if err != nil {
			return nil, err
		}
		if err := opts.SetDefaultEditLanguage(value); err != nil {
			return nil, err
		}
	}
	if c.sheetSet != nil {
		if err := opts.SetSheetSet(c.sheetSet); err != nil {
			return nil, err
		}
	}
	if c.drawObjectEventHandler != nil {
		if err := opts.SetDrawObjectEventHandler(c.drawObjectEventHandler); err != nil {
			return nil, err
		}
	}
	if c.emfRenderSetting != nil {
		value, err := enums.EmfRenderSetting(*c.emfRenderSetting)
		if err != nil {
			return nil, err
		}
		if err := opts.SetEmfRenderSetting(value); err != nil {
			return nil, err
		}
	}
	if c.customRenderSettings != nil {
		if err := opts.SetCustomRenderSettings(c.customRenderSettings); err != nil {
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
	return "xps"
}

type Option func(*Config)

func init() {
	formats.Register("xps", func() saveoptions.SaveOption {
		return New()
	})
}

// New creates a new instance of xps save options
//
// The New function creates an instance of xps SaveOption using the Functional Options Pattern. This function accepts a variable number of Option function parameters, and each Option function modifies the configuration of SaveOption.
//
// Parameters:
//
//	opts ... Option - A variable number of option functions used to configure SaveOption
//
// Return value:
// xps SaveOption - Configured instance of the saved option
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

func WithDefaultFont(value string) Option {
	return func(c *Config) {
		c.defaultFont = &value
	}
}

func WithCheckWorkbookDefaultFont(value bool) Option {
	return func(c *Config) {
		c.checkWorkbookDefaultFont = &value
	}
}

func WithCheckFontCompatibility(value bool) Option {
	return func(c *Config) {
		c.checkFontCompatibility = &value
	}
}

func WithIsFontSubstitutionCharGranularity(value bool) Option {
	return func(c *Config) {
		c.isFontSubstitutionCharGranularity = &value
	}
}

func WithOnePagePerSheet(value bool) Option {
	return func(c *Config) {
		c.onePagePerSheet = &value
	}
}

func WithAllColumnsInOnePagePerSheet(value bool) Option {
	return func(c *Config) {
		c.allColumnsInOnePagePerSheet = &value
	}
}

func WithIgnoreError(value bool) Option {
	return func(c *Config) {
		c.ignoreError = &value
	}
}

func WithOutputBlankPageWhenNothingToPrint(value bool) Option {
	return func(c *Config) {
		c.outputBlankPageWhenNothingToPrint = &value
	}
}

func WithPageIndex(value int32) Option {
	return func(c *Config) {
		c.pageIndex = &value
	}
}

func WithPageCount(value int32) Option {
	return func(c *Config) {
		c.pageCount = &value
	}
}

// WithPrintingPageType sets the printing page type: "default",
// "ignoreBlank", or "ignoreStyle".
func WithPrintingPageType(value string) Option {
	return func(c *Config) {
		c.printingPageType = &value
	}
}

// WithGridlineType sets the gridline type: "dotted" or "hair".
func WithGridlineType(value string) Option {
	return func(c *Config) {
		c.gridlineType = &value
	}
}

// WithGridlineColor sets the gridline color. Accepted forms are a Go
// color.Color (color.RGBA, color.NRGBA, color.Gray, color.Black, ...), a hex
// string ("#RRGGBB" / "#RRGGBBAA", with or without the "#"), a color name
// ("red", "Light Sea Green", matched case- and punctuation-insensitively), and
// an ARGB int. Anything else is ErrInvalidColor.
func WithGridlineColor(value interface{}) Option {
	return func(c *Config) {
		c.gridlineColor = value
	}
}

// WithTextCrossType sets the text cross type: "crossKeep", "crossOverride",
// "default", or "strictInCell".
func WithTextCrossType(value string) Option {
	return func(c *Config) {
		c.textCrossType = &value
	}
}

// WithDefaultEditLanguage sets the default edit language: "auto", "cjk", or
// "english".
func WithDefaultEditLanguage(value string) Option {
	return func(c *Config) {
		c.defaultEditLanguage = &value
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

// Disabled: this option names an engine type, which the toolkit's public API
// must not do — a caller who set it would be tied to the binding
// (docs/design.md §11). Restore it by taking the toolkit-native value instead,
// the way json.WithExportArea takes an "A1:C3" string. The original
// declaration follows verbatim.
/*
func WithDrawObjectEventHandler(value *asposecells.DrawObjectEventHandler) Option {
	return func(c *Config) {
		c.drawObjectEventHandler = value
	}
}
*/

// WithEmfRenderSetting sets the EMF render setting: "emfOnly" or
// "emfPlusPrefer".
func WithEmfRenderSetting(value string) Option {
	return func(c *Config) {
		c.emfRenderSetting = &value
	}
}

// Disabled: this option names an engine type, which the toolkit's public API
// must not do — a caller who set it would be tied to the binding
// (docs/design.md §11). Restore it by taking the toolkit-native value instead,
// the way json.WithExportArea takes an "A1:C3" string. The original
// declaration follows verbatim.
/*
func WithCustomRenderSettings(value *asposecells.CustomRenderSettings) Option {
	return func(c *Config) {
		c.customRenderSettings = value
	}
}
*/

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
