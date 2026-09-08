# Using the Converter Package for Excel File Format Conversion

Aspose.Cells for Go via C++ Toolkits provides a powerful yet intuitive `converter` package for transforming Excel files between various formats. This blog post explores the converter package's design, usage patterns, and best practices.

## Overview

The `converter` package follows a sink-based architecture that allows you to convert spreadsheets while choosing the output destination (file, in-memory bytes, or any `io.Writer`). The format is determined by the `saveoptions.SaveOption` you provide.

### Key Concepts

- **DataSource**: Where your input Excel file comes from (file, bytes, or stream)
- **SaveOption**: Determines the output format (PDF, XLSX, CSV, etc.)
- **DataSink**: Where the converted output goes (file, bytes, or writer)

## Basic Conversion

### Converting to PDF

```go
package main

import (
    "log"
    
    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/converter"
    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/pdf"
)

func main() {
    // Convert Excel to PDF with options
    err := converter.Convert(
        datasource.FilePathSource("input.xlsx"),
        pdf.New(pdf.WithOnePagePerSheet(true)),
        datasource.FilePathSink("output.pdf"),
    )
    if err != nil {
        log.Fatal(err)
    }
}
```

### Converting to Multiple Formats

The converter supports 28+ output formats. Here are some common examples:

```go
// Excel to CSV
converter.Convert(
    datasource.FilePathSource("input.xlsx"),
    csv.New(),
    datasource.FilePathSink("output.csv"),
)

// Excel to HTML
converter.Convert(
    datasource.FilePathSource("input.xlsx"),
    html.New(html.WithExportImagesAsBase64(true)),
    datasource.FilePathSink("output.html"),
)

// Excel to Markdown
converter.Convert(
    datasource.FilePathSource("input.xlsx"),
    markdown.New(markdown.WithClearData(true)),
    datasource.FilePathSink("output.md"),
)

// Excel to Image (PNG)
converter.Convert(
    datasource.FilePathSource("input.xlsx"),
    image.New(image.WithImageType("png")),
    datasource.FilePathSink("output.png"),
)
```

## Output Destinations

### Writing to Bytes

Use `BytesSink` when you need the converted data in memory:

```go
var out datasource.BytesSink
err := converter.Convert(
    datasource.FilePathSource("input.xlsx"),
    pdf.New(),
    &out,
)
if err != nil {
    log.Fatal(err)
}

// Get the bytes
pdfBytes := out.Bytes()

// Or save to file later
os.WriteFile("output.pdf", pdfBytes, 0644)
```

### Writing to a Writer

Use `WriterSink` for streaming conversion:

```go
// Convert and stream to HTTP response
func handleExcelToPDF(w http.ResponseWriter, r *http.Request) {
    var out datasource.BytesSink
    err := converter.Convert(
        datasource.NewReaderSource(r.Body),
        pdf.New(),
        &out,
    )
    if err != nil {
        http.Error(w, err.Error(), http.StatusBadRequest)
        return
    }
    
    w.Header().Set("Content-Type", "application/pdf")
    w.Header().Set("Content-Disposition", "attachment; filename=output.pdf")
    w.Write(out.Bytes())
}

// Convert and write to a buffer
var buf bytes.Buffer
err := converter.Convert(
    datasource.FilePathSource("input.xlsx"),
    csv.New(),
    datasource.NewWriterSink(&buf),
)
```

### Writing to a Folder (for Split Operations)

```go
// Save each worksheet as a separate file
opt := image.New(image.WithImageType("png"))
err := manipulator.Split(
    datasource.FilePathSource("input.xlsx"),
    opt,
    datasource.FolderSink("output/images"),
)
```

## Format Options

Each format has specific options that control the conversion behavior.

### PDF Options

```go
// Single page per sheet
pdf.New(pdf.WithOnePagePerSheet(true))

// Custom paper size
pdf.New(
    pdf.WithOnePagePerSheet(true),
    pdf.WithPageWidth(800),   // points
    pdf.WithPageHeight(1000), // points
)

// Compression
pdf.New(
    pdf.WithCompliance(pdf.PdfCompliance_Pdf15),
    pdf.WithImageQuality(80),
)
```

### HTML Options

```go
// Export as single file with base64 images
html.New(
    html.WithExportImagesAsBase64(true),
    html.WithSaveAsSingleFile(true),
)

// Custom CSS class prefix
html.New(html.WithCSSPrefix("mytable_"))
```

### Image Options

```go
// PNG conversion
image.New(image.WithImageType("png"))

// JPEG with quality
image.New(
    image.WithImageType("jpeg"),
    image.WithImageJpegQuality(90),
)

// Specify which sheets to convert
image.New(
    image.WithImageType("png"),
    image.WithSheets([]string{"Sheet1", "Sheet3"}),
)
```

### CSV Options

```go
// Standard CSV
csv.New()

// TSV (tab-separated)
csv.New(csv.WithSeparator("\t"))

// Custom encoding
csv.New(csv.WithEncoding("UTF-8"))
```

## Error Handling

The converter returns specific sentinel errors that you can classify with `errors.Is`:

```go
import (
    "errors"
    "os"
    toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
)

err := converter.Convert(source, opt, sink)
if err != nil {
    switch {
    case errors.Is(err, toolkiterrors.ErrSaveOptionNil):
        log.Println("No save option provided")
    case errors.Is(err, toolkiterrors.ErrUnsupportedFormat):
        log.Println("Unsupported output format")
    case errors.Is(err, os.ErrNotExist):
        log.Println("Input file not found")
    default:
        log.Printf("Conversion failed: %v", err)
    }
}
```

## Working with Formats Registry

The `formats` package lets you resolve formats from file extensions:

```go
import (
    "path/filepath"
    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
)

func convertByExtension(inputPath, outputPath string) error {
    // Infer format from output extension
    ext := filepath.Ext(outputPath)
    if len(ext) <= 1 {
        return fmt.Errorf("invalid output path: missing extension")
    }
    
    opt := formats.Get(ext[1:]) // "pdf" for ".pdf"
    if opt == nil {
        return fmt.Errorf("unsupported format: %s", ext[1:])
    }
    
    return converter.Convert(
        datasource.FilePathSource(inputPath),
        opt,
        datasource.FilePathSink(outputPath),
    )
}
```

## Complete Example: Batch Conversion

Here's a practical example that converts all Excel files in a directory to PDF:

```go
package main

import (
    "fmt"
    "log"
    "os"
    "path/filepath"
    
    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/converter"
    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/pdf"
)

func main() {
    inputDir := "./input"
    outputDir := "./output"
    
    // Create output directory if it doesn't exist
    if err := os.MkdirAll(outputDir, 0755); err != nil {
        log.Fatal(err)
    }
    
    // Find all Excel files
    err := filepath.Walk(inputDir, func(path string, info os.FileInfo, err error) error {
        if err != nil {
            return err
        }
        
        // Skip directories
        if info.IsDir() {
            return nil
        }
        
        // Check if it's an Excel file
        ext := filepath.Ext(path)
        if ext != ".xlsx" && ext != ".xls" && ext != ".xlsm" {
            return nil
        }
        
        // Generate output path
        relPath, _ := filepath.Rel(inputDir, path)
        relPath = filepath.ChangeExt(relPath, ".pdf")
        outputPath := filepath.Join(outputDir, relPath)
        
        // Create parent directories for output
        if err := os.MkdirAll(filepath.Dir(outputPath), 0755); err != nil {
            return err
        }
        
        // Convert
        fmt.Printf("Converting: %s -> %s\n", path, outputPath)
        err = converter.Convert(
            datasource.FilePathSource(path),
            pdf.New(pdf.WithOnePagePerSheet(true)),
            datasource.FilePathSink(outputPath),
        )
        if err != nil {
            return fmt.Errorf("convert %s: %w", path, err)
        }
        
        return nil
    })
    
    if err != nil {
        log.Fatal(err)
    }
    
    fmt.Println("All files converted successfully!")
}
```

## Performance Tips

1. **Reuse Workbook for Multiple Conversions**: If you need to convert the same workbook to multiple formats, load it once and use `datasource.BytesSource`:

```go
// Load once
data, err := os.ReadFile("input.xlsx")
if err != nil {
    log.Fatal(err)
}

// Convert to multiple formats
for _, format := range []string{"pdf", "html", "xlsx"} {
    opt := formats.Get(format)
    if opt != nil {
        var out datasource.BytesSink
        converter.Convert(
            datasource.BytesSource(data),
            opt,
            &out,
        )
        os.WriteFile("output."+format, out.Bytes(), 0644)
    }
}
```

2. **Batch Processing**: For batch operations, use worker pools to process multiple files concurrently (watch your evaluation mode limits).

3. **Memory Management**: Use `FilePathSource` and `FilePathSink` directly for large files to avoid loading entire files into memory.

## Migration from Deprecated Functions

The converter package introduced sink-based functions. The old functions are deprecated but still work:

```go
// Old way (deprecated)
// bytes, _ := converter.ConvertSpreadsheet(source, opt)
// converter.ConvertToWriter(source, w, opt)
// converter.ConvertSpreadsheetToFile(in, out)

// New way
var out datasource.BytesSink
converter.Convert(source, opt, &out)

converter.Convert(source, opt, datasource.NewWriterSink(w))

converter.Convert(
    datasource.FilePathSource(in),
    formats.Get(ext),
    datasource.FilePathSink(out),
)
```

## Conclusion

The `converter` package provides a flexible, idiomatic Go interface for Excel file format conversion. By using the sink-based architecture, you can:

- Convert to 28+ formats with a single, consistent API
- Choose your output destination (file, bytes, or writer)
- Configure format-specific options through functional options
- Handle errors gracefully with sentinel error types

For more information, check out the [full documentation](../docs/converter.md).
