# register

Package register automatically registers all supported output formats with the `formats` registry.

This is a convenience package that eliminates the need to manually register each format when you want to support all formats provided by the toolkit.

## Usage

Import the register package with a blank identifier in your main package:

```go
import (
    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/converter"
    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
    _ "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/register" // registers all formats
    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/pdf"
)
```

## What it does

The `register` package imports all format-specific saveoption subpackages, which register themselves with the `formats` registry when imported.

For example, `saveoptions/pdf` registers the "pdf" extension:

```go
// In saveoptions/pdf package
func init() {
    formats.Register("pdf", func() saveoptions.SaveOption {
        return New()
    })
}
```

When you import `_ "github.com/.../register"`, it imports all format packages, which triggers all the `init()` functions and registers all formats.

## When to use register

### Use register when:

- You want to support all output formats without manually importing each one
- You resolve formats dynamically from file extensions using `formats.Get()`
- You're writing a generic tool that processes various file types

### Don't use register when:

- You only need a specific format (e.g., only PDF) - import that format's package directly
- You want to minimize binary size (register adds all format packages to your binary)
- You want explicit control over which formats are available

## Example: Resolving format from extension

With `register` imported, you can resolve a format from a file extension:

```go
import (
    "path/filepath"
    
    _ "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/register"
    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/converter"
    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
)

func convertSpreadsheet(inputPath, outputPath string) error {
    ext := filepath.Ext(outputPath)
    if len(ext) <= 1 {
        return fmt.Errorf("invalid output path: missing extension")
    }
    
    opt := formats.Get(ext[1:]) // Get format from extension
    if opt == nil {
        return fmt.Errorf("unsupported output format: %s", ext[1:])
    }
    
    return converter.Convert(
        datasource.FilePathSource(inputPath),
        opt,
        datasource.FilePathSink(outputPath))
}
```

## Example: Without register

If you don't want to use `register`, import specific format packages:

```go
import (
    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/converter"
    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/pdf"    // only PDF
    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/xlsx" // only XLSX
)

// Use the format constructor directly - no need for formats.Get()
err := converter.Convert(
    datasource.FilePathSource("input.xlsx"),
    pdf.New(pdf.WithOnePagePerSheet(true)),
    datasource.FilePathSink("output.pdf"))
```

## Registered formats

The `register` package registers all the following formats:

- `bmp`, `csv`, `dbf`, `dif`, `docx`, `epub`, `gif`, `html`
- `jpeg`, `jpg`, `json`, `md`, `ods`, `pcl`, `pdf`, `png`
- `pptx`, `svg`, `tif`, `tiff`, `tsv`, `txt`, `xls`, `xlsb`
- `xlsm`, `xlsx`, `xltm`, `xltx`, `xml`, `xps`

## Trade-offs

| Aspect | With register | Without register |
|--------|---------------|------------------|
| Binary size | Larger (all formats included) | Smaller (only imported formats) |
| Import statements | One blank import | Multiple format imports |
| Format resolution | Dynamic via `formats.Get()` | Direct constructor call |
| IDE autocomplete | All formats available | Only imported formats |

## When not to use register

For production applications where binary size matters and you know which formats you'll use, prefer explicit imports:

```go
// Only import formats you actually use
import (
    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/pdf"
    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/xlsx"
)

// Use constructors directly
pdf.New(...)
xlsx.New(...)
```

This keeps your binary smaller and makes the supported formats explicit in your code.