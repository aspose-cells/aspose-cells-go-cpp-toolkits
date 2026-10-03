// Package image provides a SaveOption and configuration options for rendering
// spreadsheets as images (PNG, JPG, SVG, BMP, TIF/TIFF).
package image

import (
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	engine "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/engine"
	enums "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/enums"
	saveoptions "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// Config holds the native-typed option values for Image save options.
//
// Pointer fields carry presence semantics: a nil pointer means the option
// was never set (so the native default is kept), while a non-nil pointer
// means the caller explicitly requested the value, including zero/false.
// imageType is the format discriminator (returned by GetFormat and used for
// registration), so it remains a plain string rather than a pointer.
type Config struct {
	imageType string
	saveoptions.CommonConfig
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
	engine.LockEngine()
	defer engine.UnlockEngine()
	opts, err := asposecells.NewImageSaveOptions()
	if err != nil {
		return nil, err
	}

	if len(c.imageType) > 0 {
		imageOrPrintOptions, err := opts.GetImageOrPrintOptions()
		if err != nil {
			return nil, err
		}
		imageType, err := resolveImageType(c.imageType)
		if err != nil {
			return nil, err
		}
		if err := imageOrPrintOptions.SetImageType(imageType); err != nil {
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
	if c.imageType == "" {
		return "png"
	}
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
	formats.Register("jpeg", func() saveoptions.SaveOption {
		return New(WithImageType("jpeg"))
	})
	formats.Register("gif", func() saveoptions.SaveOption {
		return New(WithImageType("gif"))
	})
	formats.Register("emf", func() saveoptions.SaveOption {
		return New(WithImageType("emf"))
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

// imageTypeAliases maps the short file-extension spellings a caller is likely to
// type onto the engine's member names: the engine calls the JPEG member "Jpeg"
// and the TIFF member "Tiff", but "jpg" and "tif" name the same two formats and
// are what people actually write. The long spellings resolve without an alias.
var imageTypeAliases = map[string]string{
	"jpg": "jpeg",
	"tif": "tiff",
}

// resolveImageType resolves a requested output format name to the engine enum.
//
// An unknown name is an error rather than a fallback to the engine's
// ImageType_Unknown. That member is a sentinel for "no format recognized", not a
// format, so falling back to it would turn a mistyped option into a file written
// in whatever format the engine picks — the failure would surface far from its
// cause, if at all.
func resolveImageType(imageType string) (asposecells.ImageType, error) {
	if canonical, ok := imageTypeAliases[enums.Normalize(imageType)]; ok {
		imageType = canonical
	}
	return enums.ImageType(imageType)
}
