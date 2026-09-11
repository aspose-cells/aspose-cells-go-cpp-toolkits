// Package docx provides a SaveOption and configuration options for exporting
// spreadsheets as DOCX (Microsoft Word) documents.
package docx

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	saveoptions "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// Config holds the native-typed option values for Docx save options.
//
// Pointer fields carry presence semantics: a nil pointer means the option
// was never set (so the native default is kept), while a non-nil pointer
// means the caller explicitly requested the value, including zero/false.
type Config struct {
	saveAsEditableShapes              *bool
	embedXlsxAsChartDataSource        *bool
	asFlatOpc                         *bool
	saveElementType                   *asposecells.SaveElementType
	asNormalView                      *bool
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
	printingPageType                  *asposecells.PrintingPageType
	gridlineType                      *asposecells.GridlineType
	gridlineColor                     *asposecells.Color
	textCrossType                     *asposecells.TextCrossType
	defaultEditLanguage               *asposecells.DefaultEditLanguage
	sheetSet                          *asposecells.SheetSet
	drawObjectEventHandler            *asposecells.DrawObjectEventHandler
	emfRenderSetting                  *asposecells.EmfRenderSetting
	customRenderSettings              *asposecells.CustomRenderSettings
	saveoptions.CommonConfig
}

// Apply processes the given source byte slice as a Docx file and returns the converted output.
// This method satisfies the saveoptions.SaveOption (or equivalent) interface, enabling Docx-specific export logic.
//
// Parameters:
// - source: A byte slice representing the input spreadsheet or data source. The implementation may interpret
// this as an intermediate format (e.g., XLSX or CSV bytes) and convert it into Docx format.
//
// Returns:
// - []byte: The resulting Docx file content as a byte slice.
// - error: error information.
func (c *Config) Apply(source []byte) ([]byte, error) {
	opts, err := asposecells.NewDocxSaveOptions()
	if err != nil {
		return nil, err
	}

	if c.saveAsEditableShapes != nil {
		if err := opts.SetSaveAsEditableShapes(*c.saveAsEditableShapes); err != nil {
			return nil, err
		}
	}
	if c.embedXlsxAsChartDataSource != nil {
		if err := opts.SetEmbedXlsxAsChartDataSource(*c.embedXlsxAsChartDataSource); err != nil {
			return nil, err
		}
	}
	if c.asFlatOpc != nil {
		if err := opts.SetAsFlatOpc(*c.asFlatOpc); err != nil {
			return nil, err
		}
	}
	if c.saveElementType != nil {
		if err := opts.SetSaveElementType(*c.saveElementType); err != nil {
			return nil, err
		}
	}
	if c.asNormalView != nil {
		if err := opts.SetAsNormalView(*c.asNormalView); err != nil {
			return nil, err
		}
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
		if err := opts.SetPrintingPageType(*c.printingPageType); err != nil {
			return nil, err
		}
	}
	if c.gridlineType != nil {
		if err := opts.SetGridlineType(*c.gridlineType); err != nil {
			return nil, err
		}
	}
	if c.gridlineColor != nil {
		if err := opts.SetGridlineColor(c.gridlineColor); err != nil {
			return nil, err
		}
	}
	if c.textCrossType != nil {
		if err := opts.SetTextCrossType(*c.textCrossType); err != nil {
			return nil, err
		}
	}
	if c.defaultEditLanguage != nil {
		if err := opts.SetDefaultEditLanguage(*c.defaultEditLanguage); err != nil {
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
		if err := opts.SetEmfRenderSetting(*c.emfRenderSetting); err != nil {
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
	return "docx"
}

type Option func(*Config)

func init() {
	formats.Register("docx", func() saveoptions.SaveOption {
		return New()
	})
}

// New creates a new instance of docx save options
//
// The New function creates an instance of docx SaveOption using the Functional Options Pattern. This function accepts a variable number of Option function parameters, and each Option function modifies the configuration of SaveOption.
//
// Parameters:
//
//	opts ... Option - A variable number of option functions used to configure SaveOption
//
// Return value:
// docx SaveOption - Configured instance of the saved option
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

func WithSaveAsEditableShapes(value bool) Option {
	return func(c *Config) {
		c.saveAsEditableShapes = &value
	}
}

func WithEmbedXlsxAsChartDataSource(value bool) Option {
	return func(c *Config) {
		c.embedXlsxAsChartDataSource = &value
	}
}

func WithAsFlatOpc(value bool) Option {
	return func(c *Config) {
		c.asFlatOpc = &value
	}
}

func WithSaveElementType(value asposecells.SaveElementType) Option {
	return func(c *Config) {
		c.saveElementType = &value
	}
}
func WithAsNormalView(value bool) Option {
	return func(c *Config) {
		c.asNormalView = &value
	}
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

func WithPrintingPageType(value asposecells.PrintingPageType) Option {
	return func(c *Config) {
		c.printingPageType = &value
	}
}
func WithGridlineType(value asposecells.GridlineType) Option {
	return func(c *Config) {
		c.gridlineType = &value
	}
}
func WithGridlineColor(value *asposecells.Color) Option {
	return func(c *Config) {
		c.gridlineColor = value
	}
}
func WithTextCrossType(value asposecells.TextCrossType) Option {
	return func(c *Config) {
		c.textCrossType = &value
	}
}
func WithDefaultEditLanguage(value asposecells.DefaultEditLanguage) Option {
	return func(c *Config) {
		c.defaultEditLanguage = &value
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
		c.emfRenderSetting = &value
	}
}
func WithCustomRenderSettings(value *asposecells.CustomRenderSettings) Option {
	return func(c *Config) {
		c.customRenderSettings = value
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
