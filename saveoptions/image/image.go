package image

import (
	. "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
	"strconv"
	"strings"
)

type Config struct {
	imageType                 string
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

// Apply processes the given source byte slice as an image file and returns the converted output.
// This method satisfies the saveoptions.SaveOption (or equivalent) interface, enabling Image-specific export logic.
//
// Parameters:
// - source: A byte slice representing the input spreadsheet or data source. The implementation may interpret
// this as an intermediate format (e.g., XLSX or CSV bytes) and convert it into image format(png,jpg,img,svg, and so on).
//
// Returns:
// - []byte: The resulting image file content as a byte slice.
// - error: error information.
func (c *Config) Apply(source []byte) ([]byte, error) {
	opts, _ := asposecells.NewImageSaveOptions()

	if len(c.imageType) > 0 {
		imageOrPrintOptions, _ := opts.GetImageOrPrintOptions()
		imageOrPrintOptions.SetImageType(toImageType(c.imageType))
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

func WithImageType(value string) Option {
	return func(c *Config) {
		c.imageType = value
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

var imageTypeRegistry = make(map[string]asposecells.ImageType)

func init() {
	imageTypeRegistry["png"] = asposecells.ImageType_Png
	imageTypeRegistry["emf"] = asposecells.ImageType_Emf
	imageTypeRegistry["wmf"] = asposecells.ImageType_Wmf
	imageTypeRegistry["pict"] = asposecells.ImageType_Pict
	imageTypeRegistry["jpg"] = asposecells.ImageType_Jpeg
	imageTypeRegistry["jpeg"] = asposecells.ImageType_Jpeg
	imageTypeRegistry["bmp"] = asposecells.ImageType_Bmp
	imageTypeRegistry["gif"] = asposecells.ImageType_Gif
	imageTypeRegistry["tif"] = asposecells.ImageType_Tiff
	imageTypeRegistry["tiff"] = asposecells.ImageType_Tiff
	imageTypeRegistry["svg"] = asposecells.ImageType_Svg
	imageTypeRegistry["svm"] = asposecells.ImageType_Svm
	imageTypeRegistry["gltf"] = asposecells.ImageType_Gltf
	imageTypeRegistry["emf"] = asposecells.ImageType_OfficeCompatibleEmf
	imageTypeRegistry["webp"] = asposecells.ImageType_WebP
}
func toImageType(imageType string) asposecells.ImageType {
	value := strings.ToLower(imageType)
	if val, ok := imageTypeRegistry[value]; ok {
		return val
	}
	return asposecells.ImageType_Unknown
}
