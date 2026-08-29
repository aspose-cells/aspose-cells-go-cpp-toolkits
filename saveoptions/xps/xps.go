package xps

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
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
	printingPageType                  *asposecells.PrintingPageType
	gridlineType                      *asposecells.GridlineType
	gridlineColor                     *asposecells.Color
	textCrossType                     *asposecells.TextCrossType
	defaultEditLanguage               *asposecells.DefaultEditLanguage
	sheetSet                          *asposecells.SheetSet
	drawObjectEventHandler            *asposecells.DrawObjectEventHandler
	emfRenderSetting                  *asposecells.EmfRenderSetting
	customRenderSettings              *asposecells.CustomRenderSettings
	clearData                         *bool
	cachedFileFolder                  *string
	validateMergedAreas               *bool
	mergeAreas                        *bool
	createDirectory                   *bool
	sortNames                         *bool
	sortExternalNames                 *bool
	refreshChartCache                 *bool
	checkExcelRestriction             *bool
	updateSmartArt                    *bool
	encryptDocumentProperties         *bool
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
