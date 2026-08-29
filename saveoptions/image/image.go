// Package image provides a SaveOption and configuration options for rendering
// spreadsheets as images (PNG, JPG, SVG, BMP, TIF/TIFF).
package image

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	saveoptions "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
	"strings"
)

// Config holds the native-typed option values for Image save options.
//
// Pointer fields carry presence semantics: a nil pointer means the option
// was never set (so the native default is kept), while a non-nil pointer
// means the caller explicitly requested the value, including zero/false.
// imageType is the format discriminator (returned by GetFormat and used for
// registration), so it remains a plain string rather than a pointer.
type Config struct {
	imageType                 string
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
	opts, err := asposecells.NewImageSaveOptions()
	if err != nil {
		return nil, err
	}

	if len(c.imageType) > 0 {
		imageOrPrintOptions, err := opts.GetImageOrPrintOptions()
		if err != nil {
			return nil, err
		}
		if err := imageOrPrintOptions.SetImageType(toImageType(c.imageType)); err != nil {
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
	return c.imageType
}

type Option func(*Config)

func init() {
	formats.Register("png", func() saveoptions.SaveOption {
		return New(WithImageType("png"))
	})
	formats.Register("jpg", func() saveoptions.SaveOption {
		return New(WithImageType("jpg"))
	})
	formats.Register("svg", func() saveoptions.SaveOption {
		return New(WithImageType("svg"))
	})
	formats.Register("bmp", func() saveoptions.SaveOption {
		return New(WithImageType("bmp"))
	})
	formats.Register("tif", func() saveoptions.SaveOption {
		return New(WithImageType("tif"))
	})
	formats.Register("tiff", func() saveoptions.SaveOption {
		return New(WithImageType("tiff"))
	})

}

// New creates a new instance of image save options
//
// The New function creates an instance of image SaveOption using the Functional Options Pattern. This function accepts a variable number of Option function parameters, and each Option function modifies the configuration of SaveOption.
//
// Parameters:
//
//	opts ... Option - A variable number of option functions used to configure SaveOption
//
// Return value:
// image SaveOption - Configured instance of the saved option
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

func WithImageType(value string) Option {
	return func(c *Config) {
		c.imageType = value
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
