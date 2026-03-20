package ooxml

import (
	. "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
	"strconv"
)

type Config struct {
	asFlatOpc                 string
	exportCellName            string
	updateZoom                string
	enableZip64               string
	embedOoxmlAsOleObject     string
	compressionType           string
	clearData                 string
	cachedFileFolder          string
	validateMergedAreas       string
	mergeAreas                string
	createDirectory           string
	sortNames                 string
	sortExternalNames         string
	refreshChartCache         string
	checkExcelRestriction     string
	updateSmartArt            string
	encryptDocumentProperties string
}

func (c *Config) Apply(source []byte) ([]byte, error) {
	opts, _ := asposecells.NewOoxmlSaveOptions()

	if len(c.asFlatOpc) > 0 {
		if v, err := strconv.ParseBool(c.asFlatOpc); err == nil {
			opts.SetAsFlatOpc(v)
		}
	}
	if len(c.exportCellName) > 0 {
		if v, err := strconv.ParseBool(c.exportCellName); err == nil {
			opts.SetExportCellName(v)
		}
	}
	if len(c.updateZoom) > 0 {
		if v, err := strconv.ParseBool(c.updateZoom); err == nil {
			opts.SetUpdateZoom(v)
		}
	}
	if len(c.enableZip64) > 0 {
		if v, err := strconv.ParseBool(c.enableZip64); err == nil {
			opts.SetEnableZip64(v)
		}
	}
	if len(c.embedOoxmlAsOleObject) > 0 {
		if v, err := strconv.ParseBool(c.embedOoxmlAsOleObject); err == nil {
			opts.SetEmbedOoxmlAsOleObject(v)
		}
	}
	if v, err := strconv.ParseInt(c.compressionType, 10, 32); err == nil {
		if vv, err2 := asposecells.Int32ToOoxmlCompressionType(int32(v)); err2 == nil {
			opts.SetCompressionType(vv)
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

func WithAsFlatOpc(value bool) Option {
	return func(c *Config) {
		c.asFlatOpc = strconv.FormatBool(value)
	}
}

func WithExportCellName(value bool) Option {
	return func(c *Config) {
		c.exportCellName = strconv.FormatBool(value)
	}
}

func WithUpdateZoom(value bool) Option {
	return func(c *Config) {
		c.updateZoom = strconv.FormatBool(value)
	}
}

func WithEnableZip64(value bool) Option {
	return func(c *Config) {
		c.enableZip64 = strconv.FormatBool(value)
	}
}

func WithEmbedOoxmlAsOleObject(value bool) Option {
	return func(c *Config) {
		c.embedOoxmlAsOleObject = strconv.FormatBool(value)
	}
}

func WithCompressionType(value asposecells.OoxmlCompressionType) Option {
	return func(c *Config) {
		c.compressionType = strconv.FormatInt(int64(value), 10)
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
