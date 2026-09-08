# formats

Package formats maps file extensions and Aspose.Cells format types to `saveoptions.SaveOption` factories.

Implementations register themselves by extension so callers can resolve an output format from a filename or from a `FileFormatType`. The registry is concurrency-safe: `Register`, `Unregister`, `Get`, and `List` may be called from multiple goroutines, and `List` returns a sorted, stable snapshot.

## Overview

The `formats` package serves as a registry that maps file extensions (like "xlsx", "pdf", "csv") to factory functions that create the corresponding `saveoptions.SaveOption`. This enables format selection based on file extension inference.

## Functions

### Get

```go
func Get(ext string) saveoptions.SaveOption
```

Returns a new `SaveOption` for the given extension, or `nil` if the extension is not registered. The extension lookup is case-insensitive.

Example:

```go
// Get a PDF save option from extension
opt := formats.Get("pdf")
if opt == nil {
    log.Fatal("PDF format not registered")
}

// Get an XLSX save option from extension
opt := formats.Get("XLSX") // case-insensitive, same as "xlsx"
```

### List

```go
func List() []string
```

Returns all registered extensions in sorted order. Safe for concurrent use.

Example:

```go
exts := formats.List()
// exts might be: ["bmp", "csv", "dbf", "dif", "docx", "epub", ...]
```

### Register

```go
func Register(ext string, factory func() saveoptions.SaveOption)
```

Maps a file extension to a factory that produces its `SaveOption`. The extension is normalized to lower case and lookups are case-insensitive. `Register` is safe for concurrent use and may be called at any time.

> **Note**: Registering a format that was already registered replaces the previous factory.

### Unregister

```go
func Unregister(ext string)
```

Removes a previously registered extension. It is safe for concurrent use and is a no-op if the extension is not registered.

## Using formats with converter

The most common use case is to convert a spreadsheet to a format inferred from the output file extension:

```go
inputPath := "input.xlsx"
outputPath := "output.pdf"

// Infer format from output extension
ext := filepath.Ext(outputPath) // ".pdf"
opt := formats.Get(ext[1:])     // "pdf" -> PDF save option

if opt == nil {
    log.Fatalf("unsupported output format: %s", ext[1:])
}

err := converter.Convert(
    datasource.FilePathSource(inputPath),
    opt,
    datasource.FilePathSink(outputPath))
```

## Supported formats

The following formats are registered by default (via the `register` package or manual registration):

| Extension | Format | Description |
|-----------|--------|-------------|
| `bmp` | BMP | Bitmap Image Format |
| `csv` | CSV | Comma Separated Values |
| `dbf` | DBF | Database File Format |
| `dif` | DIF | Data Interchange Format |
| `docx` | DOCX | Microsoft Word Document |
| `epub` | EPUB | Electronic Publication |
| `gif` | GIF | Graphics Interchange Format |
| `html` | HTML | HyperText Markup Language |
| `jpeg` / `jpg` | JPEG | Joint Photographic Experts Group |
| `json` | JSON | JavaScript Object Notation |
| `md` | Markdown | Markdown text format |
| `ods` | ODS | OpenDocument Spreadsheet |
| `pcl` | PCL | Printer Command Language |
| `pdf` | PDF | Portable Document Format |
| `png` | PNG | Portable Network Graphics |
| `pptx` | PPTX | PowerPoint Document |
| `svg` | SVG | Scalable Vector Graphics |
| `tif` / `tiff` | TIFF | Tagged Image File Format |
| `tsv` | TSV | Tab Separated Values |
| `txt` | TXT | Plain Text |
| `xls` | XLS | Excel 97-2003 Workbook |
| `xlsb` | XLSB | Excel Binary Workbook |
| `xlsm` | XLSM | Excel Macro-Enabled Workbook |
| `xlsx` | XLSX | Excel Workbook |
| `xltm` | XLTM | Excel Macro-Enabled Template |
| `xltx` | XLTX | Excel Template |
| `xml` | XML | Extensible Markup Language |
| `xps` | XPS | XML Paper Specification |

## Custom formats

You can register custom or third-party `SaveOption` implementations:

```go
// Define a custom save option
type CustomOption struct {
    quality int
}

func (o *CustomOption) Apply(data []byte) ([]byte, error) {
    // custom conversion logic
    return data, nil
}

func (o *CustomOption) GetFormat() string {
    return "custom"
}

// Register it with an extension
formats.Register("custom", func() saveoptions.SaveOption {
    return &CustomOption{quality: 90}
})

// Now use it
opt := formats.Get("custom")
```

## Relationship with saveoptions

- **saveoptions**: Defines the `SaveOption` interface and format-specific option types (e.g., `pdf.SaveOption`, `csv.SaveOption`)
- **formats**: Provides a registry that maps extensions to factory functions that create those option types

The `formats` package doesn't define formats itself; it only maps extensions to factories that create `saveoptions.SaveOption` implementations.