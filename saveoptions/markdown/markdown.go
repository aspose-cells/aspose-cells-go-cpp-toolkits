// Package markdown provides a SaveOption and configuration options for
// exporting spreadsheets as Markdown tables.
package markdown

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	engine "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/engine"
	enums "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/enums"
	saveoptions "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// Config holds the native-typed option values for Markdown save options.
//
// Pointer fields carry presence semantics: a nil pointer means the option
// was never set (so the native default is kept), while a non-nil pointer
// means the caller explicitly requested the value, including zero/false.
type Config struct {
	encoding                   *string
	formatStrategy             *string
	lineSeparator              *string
	tableHeaderType            *string
	sheetSet                   *asposecells.SheetSet
	exportImagesAsBase64       *bool
	calculateFormula           *bool
	exportHyperlinkAsReference *bool
	alignColumnPadding         *byte
	splitTablesByBlankRow      *bool
	officeMathOutputType       *string
	saveoptions.CommonConfig
}

// Apply processes the given source byte slice as a Markdown file and returns the converted output.
// This method satisfies the saveoptions.SaveOption (or equivalent) interface, enabling Markdown-specific export logic.
//
// Parameters:
// - source: A byte slice representing the input spreadsheet or data source. The implementation may interpret
// this as an intermediate format (e.g., XLSX or CSV bytes) and convert it into Markdown format.
//
// Returns:
// - []byte: The resulting Markdown file content as a byte slice.
// - error: error information.
func (c *Config) Apply(source []byte) ([]byte, error) {
	engine.LockEngine()
	defer engine.UnlockEngine()
	opts, err := asposecells.NewMarkdownSaveOptions()
	if err != nil {
		return nil, err
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
	if c.formatStrategy != nil {
		value, err := enums.CellValueFormatStrategy(*c.formatStrategy)
		if err != nil {
			return nil, err
		}
		if err := opts.SetFormatStrategy(value); err != nil {
			return nil, err
		}
	}
	if c.lineSeparator != nil {
		if err := opts.SetLineSeparator(*c.lineSeparator); err != nil {
			return nil, err
		}
	}
	if c.tableHeaderType != nil {
		value, err := enums.MarkdownTableHeaderType(*c.tableHeaderType)
		if err != nil {
			return nil, err
		}
		if err := opts.SetTableHeaderType(value); err != nil {
			return nil, err
		}
	}
	if c.sheetSet != nil {
		if err := opts.SetSheetSet(c.sheetSet); err != nil {
			return nil, err
		}
	}
	if c.exportImagesAsBase64 != nil {
		if err := opts.SetExportImagesAsBase64(*c.exportImagesAsBase64); err != nil {
			return nil, err
		}
	}
	if c.calculateFormula != nil {
		if err := opts.SetCalculateFormula(*c.calculateFormula); err != nil {
			return nil, err
		}
	}
	if c.exportHyperlinkAsReference != nil {
		if err := opts.SetExportHyperlinkAsReference(*c.exportHyperlinkAsReference); err != nil {
			return nil, err
		}
	}
	if c.alignColumnPadding != nil {
		if err := opts.SetAlignColumnPadding(*c.alignColumnPadding); err != nil {
			return nil, err
		}
	}
	if c.splitTablesByBlankRow != nil {
		if err := opts.SetSplitTablesByBlankRow(*c.splitTablesByBlankRow); err != nil {
			return nil, err
		}
	}
	if c.officeMathOutputType != nil {
		value, err := enums.HtmlOfficeMathOutputType(*c.officeMathOutputType)
		if err != nil {
			return nil, err
		}
		if err := opts.SetOfficeMathOutputType(value); err != nil {
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
	return "md"
}

type Option func(*Config)

func init() {
	formats.Register("md", func() saveoptions.SaveOption {
		return New()
	})
}

// New creates a new instance of md save options
//
// The New function creates an instance of md SaveOption using the Functional Options Pattern. This function accepts a variable number of Option function parameters, and each Option function modifies the configuration of SaveOption.
//
// Parameters:
//
//	opts ... Option - A variable number of option functions used to configure SaveOption
//
// Return value:
// md SaveOption - Configured instance of the saved option
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

// WithEncoding sets the encoding: "ascii", "default", "unicode", or "utf8".
func WithEncoding(value string) Option {
	return func(c *Config) {
		c.encoding = &value
	}
}

// WithFormatStrategy sets the format strategy: "cellStyle", "displayString",
// "displayStyle", or "none".
func WithFormatStrategy(value string) Option {
	return func(c *Config) {
		c.formatStrategy = &value
	}
}
func WithLineSeparator(value string) Option {
	return func(c *Config) {
		c.lineSeparator = &value
	}
}

// WithTableHeaderType sets the table header type: "columnHeader", "empty",
// or "firstRow".
func WithTableHeaderType(value string) Option {
	return func(c *Config) {
		c.tableHeaderType = &value
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

func WithExportImagesAsBase64(value bool) Option {
	return func(c *Config) {
		c.exportImagesAsBase64 = &value
	}
}

func WithCalculateFormula(value bool) Option {
	return func(c *Config) {
		c.calculateFormula = &value
	}
}

func WithExportHyperlinkAsReference(value bool) Option {
	return func(c *Config) {
		c.exportHyperlinkAsReference = &value
	}
}

func WithAlignColumnPadding(value byte) Option {
	return func(c *Config) {
		c.alignColumnPadding = &value
	}
}

func WithSplitTablesByBlankRow(value bool) Option {
	return func(c *Config) {
		c.splitTablesByBlankRow = &value
	}
}

// WithOfficeMathOutputType sets the office math output type: "image" or
// "mathML".
func WithOfficeMathOutputType(value string) Option {
	return func(c *Config) {
		c.officeMathOutputType = &value
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
