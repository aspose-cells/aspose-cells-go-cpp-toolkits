// Package sqlScript provides a SaveOption and configuration options for
// exporting spreadsheets as SQL scripts.
package sqlscript

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	saveoptions "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// Config holds the native-typed option values for Sql save options.
//
// Pointer fields carry presence semantics: a nil pointer means the option
// was never set (so the native default is kept), while a non-nil pointer
// means the caller explicitly requested the value, including zero/false.
type Config struct {
	checkIfTableExists        *bool
	columnTypeMap             *asposecells.SqlScriptColumnTypeMap
	checkAllDataForColumnType *bool
	addBlankLineBetweenRows   *bool
	separator                 *byte
	operatorType              *asposecells.SqlScriptOperatorType
	primaryKey                *int32
	createTable               *bool
	idName                    *string
	startId                   *int32
	tableName                 *string
	exportAsString            *bool
	sheetIndexes              []int32
	exportArea                *asposecells.CellArea
	hasHeaderRow              *bool
	clearData                 *bool
	cachedFileFolder          *string
	validateMergedAreas       *bool
	mergeAreas                *bool
	createDirectory           *bool
	sortNames                 *bool
	sortExternalNames         *bool
	refreshChartCache         *bool
	checkExcelRestriction     *bool
	updateSmartArt            *bool
	encryptDocumentProperties *bool
}

// Apply processes the given source byte slice as a sql file and returns the converted output.
// This method satisfies the saveoptions.SaveOption (or equivalent) interface, enabling Sql-specific export logic.
//
// Parameters:
// - source: A byte slice representing the input spreadsheet or data source. The implementation may interpret
// this as an intermediate format (e.g., XLSX or CSV bytes) and convert it into Sql format.
//
// Returns:
// - []byte: The resulting Sql file content as a byte slice.
// - error: error information.
func (c *Config) Apply(source []byte) ([]byte, error) {
	opts, err := asposecells.NewSqlScriptSaveOptions()
	if err != nil {
		return nil, err
	}

	if c.checkIfTableExists != nil {
		if err := opts.SetCheckIfTableExists(*c.checkIfTableExists); err != nil {
			return nil, err
		}
	}
	if c.columnTypeMap != nil {
		if err := opts.SetColumnTypeMap(c.columnTypeMap); err != nil {
			return nil, err
		}
	}
	if c.checkAllDataForColumnType != nil {
		if err := opts.SetCheckAllDataForColumnType(*c.checkAllDataForColumnType); err != nil {
			return nil, err
		}
	}
	if c.addBlankLineBetweenRows != nil {
		if err := opts.SetAddBlankLineBetweenRows(*c.addBlankLineBetweenRows); err != nil {
			return nil, err
		}
	}
	if c.separator != nil {
		if err := opts.SetSeparator(*c.separator); err != nil {
			return nil, err
		}
	}
	if c.operatorType != nil {
		if err := opts.SetOperatorType(*c.operatorType); err != nil {
			return nil, err
		}
	}
	if c.primaryKey != nil {
		if err := opts.SetPrimaryKey(*c.primaryKey); err != nil {
			return nil, err
		}
	}
	if c.createTable != nil {
		if err := opts.SetCreateTable(*c.createTable); err != nil {
			return nil, err
		}
	}
	if c.idName != nil {
		if err := opts.SetIdName(*c.idName); err != nil {
			return nil, err
		}
	}
	if c.startId != nil {
		if err := opts.SetStartId(*c.startId); err != nil {
			return nil, err
		}
	}
	if c.tableName != nil {
		if err := opts.SetTableName(*c.tableName); err != nil {
			return nil, err
		}
	}
	if c.exportAsString != nil {
		if err := opts.SetExportAsString(*c.exportAsString); err != nil {
			return nil, err
		}
	}
	if c.sheetIndexes != nil {
		if err := opts.SetSheetIndexes(c.sheetIndexes); err != nil {
			return nil, err
		}
	}
	if c.exportArea != nil {
		if err := opts.SetExportArea(c.exportArea); err != nil {
			return nil, err
		}
	}
	if c.hasHeaderRow != nil {
		if err := opts.SetHasHeaderRow(*c.hasHeaderRow); err != nil {
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
	return "sql"
}

type Option func(*Config)

func init() {
	formats.Register("sql", func() saveoptions.SaveOption {
		return New()
	})
}

// New creates a new instance of sql save options
//
// The New function creates an instance of sql SaveOption using the Functional Options Pattern. This function accepts a variable number of Option function parameters, and each Option function modifies the configuration of SaveOption.
//
// Parameters:
//
//	opts ... Option - A variable number of option functions used to configure SaveOption
//
// Return value:
// sql SaveOption - Configured instance of the saved option
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

func WithCheckIfTableExists(value bool) Option {
	return func(c *Config) {
		c.checkIfTableExists = &value
	}
}

func WithColumnTypeMap(value *asposecells.SqlScriptColumnTypeMap) Option {
	return func(c *Config) {
		c.columnTypeMap = value
	}
}
func WithCheckAllDataForColumnType(value bool) Option {
	return func(c *Config) {
		c.checkAllDataForColumnType = &value
	}
}

func WithAddBlankLineBetweenRows(value bool) Option {
	return func(c *Config) {
		c.addBlankLineBetweenRows = &value
	}
}

func WithSeparator(value byte) Option {
	return func(c *Config) {
		c.separator = &value
	}
}

func WithOperatorType(value asposecells.SqlScriptOperatorType) Option {
	return func(c *Config) {
		c.operatorType = &value
	}
}
func WithPrimaryKey(value int32) Option {
	return func(c *Config) {
		c.primaryKey = &value
	}
}

func WithCreateTable(value bool) Option {
	return func(c *Config) {
		c.createTable = &value
	}
}

func WithIdName(value string) Option {
	return func(c *Config) {
		c.idName = &value
	}
}

func WithStartId(value int32) Option {
	return func(c *Config) {
		c.startId = &value
	}
}

func WithTableName(value string) Option {
	return func(c *Config) {
		c.tableName = &value
	}
}

func WithExportAsString(value bool) Option {
	return func(c *Config) {
		c.exportAsString = &value
	}
}

func WithSheetIndexes(value []int32) Option {
	return func(c *Config) {
		c.sheetIndexes = value
	}
}
func WithExportArea(value *asposecells.CellArea) Option {
	return func(c *Config) {
		c.exportArea = value
	}
}
func WithHasHeaderRow(value bool) Option {
	return func(c *Config) {
		c.hasHeaderRow = &value
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
