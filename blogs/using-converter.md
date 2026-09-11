---
title: "Converter: one line at its simplest, any environment at its most flexible"
description: "Convert Excel files with one Go function. converter.Convert(source, saveOption, sink) turns xlsx into PDF, CSV, HTML, Markdown and 32 total formats — from a concise file-to-file one-liner to fully flexible memory, stream and batch environments."
keywords:
  - "Excel to PDF Go"
  - "Aspose.Cells Go converter"
  - "xlsx to csv Go"
  - "converter.Convert"
  - "Excel format conversion Go"
  - "Aspose.Cells for Go via C++"
  - "formats.Get"
  - "Excel conversion toolkit"
---
# Converter: one line at its simplest, any environment at its most flexible

The `converter` package of Aspose.Cells for Go via C++ Toolkits reduces every "turn a spreadsheet into another format" need to **one entry function**:

```go
converter.Convert(source, saveOption, sink) error
```

It has only three parameters, and each answers exactly one question:

| Parameter | Type | Question it answers | Typical values |
| --- | --- | --- | --- |
| `source` | `datasource.DataSource` | Where does the input come from? | `FilePathSource(path)`, `BytesSource(data)`, `NewReaderSource(r)` |
| `saveOption` | `saveoptions.SaveOption` | Output as what, tuned how? | `formats.Get("pdf")` for defaults by name, or `pdf.New(pdf.WithOnePagePerSheet(true))` for fine control |
| `sink` | `datasource.DataSink` | Where does the output go? | `FilePathSink(path)`, `&BytesSink{}`, `NewWriterSink(w)` |

For the **most concise** style, take the default on all three axes; for the **most flexible**, turn each knob independently. Because the three axes are independent, the same `Convert` naturally fits file, memory, HTTP, and batch environments. (There is also a legacy one-liner, `ConvertSpreadsheetToFile`, for the plain file→file case — compared in section 1 below.) The rest of this article moves from most concise to most flexible, growing the previous snippet each step rather than switching to a different API.

---

## 1. The simplest: one call for one format

### 1.1 A built-in file→file one-liner: `ConvertSpreadsheetToFile`

The package also ships an older, even shorter entry for the classic "convert this file into that file" case. All it takes is two paths:

```go
package main

import (
	"log"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/converter"
	_ "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/register" // registers every output format
)

func main() {
	// Both ends are files; the output format is inferred from the extension.
	if err := converter.ConvertSpreadsheetToFile("input.xlsx", "output.pdf"); err != nil {
		log.Fatal(err)
	}
}
```

> **Deprecated.** `ConvertSpreadsheetToFile` is a thin wrapper kept for backward compatibility; new code is directed to `converter.Convert` below. For the file→file case the two are exactly equivalent — this is all `ConvertSpreadsheetToFile` does:

```go
opt := formats.Get(strings.TrimPrefix(filepath.Ext("output.pdf"), ".")) // "pdf"
converter.Convert(
	datasource.FilePathSource("input.xlsx"), // input: a file
	opt,                                     // format: taken from the extension
	datasource.FilePathSink("output.pdf"),   // output: a file
)
```

`ConvertSpreadsheetToFile` and the `Convert` form above produce identical results; what differs is the scenario, as compared next.

### 1.2 The recommended one-liner: `Convert` with defaults

```go
package main

import (
	"log"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/converter"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/pdf"
)

func main() {
	if err := converter.Convert(
		datasource.FilePathSource("input.xlsx"), // input: a file
		pdf.New(),                               // format: PDF, all defaults
		datasource.FilePathSink("output.pdf"),   // output: a file
	); err != nil {
		log.Fatal(err)
	}
}
```

All three knobs take their default: input = a file, format = `pdf.New()` with zero configuration, output = a file. That is the "most concise" shape — no manual `os.Open`/`os.Create`, no object allocation, no handle closing; the source/sink manage resources for you.

### 1.3 The two file→file shapes compared

| | `ConvertSpreadsheetToFile(in, out)` | `converter.Convert(source, opt, sink)` |
| --- | --- | --- |
| Input | a file path only | any `DataSource`: path, `[]byte`, or reader |
| Output | a file path only | any `DataSink`: path, bytes, or `io.Writer` |
| Output format | always inferred from the extension, defaults only | any `saveOption`: default or format-specific tuning |
| Extra import | needs the format registered (`_ "…/register"`, or the matching `saveoptions/*` subpackage) | none, when you pass a constructor like `pdf.New()` |
| Status | deprecated thin wrapper (stable, kept for compatibility) | the documented single entry point for new code |

Same file→file outcome, two different scenarios:

- **`ConvertSpreadsheetToFile`** — you literally have two paths and are happy with the engine's default settings for the target format. Ideal for one-off conversion scripts, migration batches, and test helpers. Its price: the input, the output, and the settings can never leave the "file path + defaults" shape — and it is deprecated.
- **`converter.Convert`** — you want the same one-liner, but any of the three arguments stays replaceable: tomorrow the input is an upload's `[]byte`, the target needs `pdf.WithOnePagePerSheet(true)`, or the output should go straight into an HTTP response. Those are exactly the scenarios of sections 3 and 4 — write it once with `Convert`, and each change is a one-argument edit.

## 2. The simplest × N: one function for all 32 formats

The next level of conciseness is using the **same function** no matter which format you target. Because `saveOption` can be computed from the output file's extension, pulling that inference out yields a "universal file converter":

```go
package main

import (
	"errors"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/converter"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	_ "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/register" // registers every output format
)

// convert re-saves in as out; the output format is read from out's extension.
// The same code handles every built-in format (pdf/csv/html/md/xlsx/png/…).
// It is exactly what the deprecated ConvertSpreadsheetToFile does internally.
func convert(in, out string) error {
	opt := formats.Get(strings.TrimPrefix(filepath.Ext(out), ".")) // ".pdf" -> "pdf"
	if opt == nil {
		return fmt.Errorf("%w: %q", toolkiterrors.ErrUnsupportedFormat, filepath.Ext(out))
	}
	return converter.Convert(
		datasource.FilePathSource(in), // where the input comes from
		opt,                           // what to produce: decided by the extension
		datasource.FilePathSink(out),  // where the output goes
	)
}

func main() {
	if len(os.Args) != 3 {
		log.Fatalf("usage: %s <input.xlsx> <output.ext>", os.Args[0])
	}
	if err := convert(os.Args[1], os.Args[2]); err != nil {
		if errors.Is(err, toolkiterrors.ErrUnsupportedFormat) {
			log.Printf("no registered output format for %q", filepath.Ext(os.Args[2]))
		}
		log.Fatal(err)
	}
	fmt.Println("done")
}
```

Usage stays identical — only the output extension changes:

```console
$ go run . input.xlsx report.pdf     # to PDF
$ go run . input.xlsx report.csv     # to CSV
$ go run . input.xlsx report.html    # to HTML
$ go run . input.xlsx notes.md       # to Markdown
$ go run . input.xlsx sheet.png      # to a PNG image
$ go run . input.xlsx legacy.xls     # to legacy Excel
```

### What is that blank `_ "…/register"` import?

There are two ways to obtain a `saveOption`; this is the only concept the "most concise" style needs to understand:

- **Look it up from the registry by format name** (`formats.Get`): the extension must be present in the `formats` registry. Each `saveoptions/*` subpackage registers itself in its own `init()`, but only when imported. Import the subpackages you need one by one (only `pdf`, say, if you convert to PDF), or **blank-import `…/register` once**, which registers all 32 formats at once — the "universal converter" above relies on that single line.
- **Call a constructor directly** (`pdf.New()`, `markdown.New()`): bypasses the registry, no `register` import needed; the cost is one import line per format you use.

In short: only use `_ "…/register"` when you resolve via `formats.Get` + the output extension; passing a constructor directly needs nothing extra. The blank import is harmless and is always present in the official examples.

## 3. The most flexible: turn each knob independently

When the need goes beyond "file to file with default settings", you do not change API — the three `Convert` parameters are independently swappable.

### 1. Input (source): from a file, to memory, to any stream

```go
datasource.FilePathSource("input.xlsx") // a file on disk
datasource.BytesSource(data)            // []byte already in memory (downloads, cached reuse)
datasource.NewReaderSource(r.Body)      // any io.Reader (HTTP body, compressed stream, …)
```

> Note: the underlying engine needs random access to the whole content, so `ReaderSource` buffers the stream fully on first access; every later `Open` serves a copy of that buffer. Use it confidently for HTTP, just know it is not a true "convert while streaming".

### 2. Format (saveOption): defaults, or tuned

| Scenario | Writing style | Why |
| --- | --- | --- |
| Some format with default settings | `formats.Get("pdf")` | looked up by extension; case-insensitive; `nil` if unregistered |
| A format with format-specific options | `pdf.New(pdf.WithOnePagePerSheet(true))` | functional options |
| Default settings, explicit format | `ooxml.New()` / `csv.New()` / … | one zero-argument constructor per format |
| A custom private format | `formats.Register("fmt", factory)` | plug into the registry — see [docs/register.md](../docs/register.md) |

Functional options appear only on the formats that have them: one page per worksheet in PDF is `WithOnePagePerSheet`; images inline in HTML is `WithExportImagesAsBase64`; TSV is `formats.Get("tsv")`. Options are pointer-semantic: omit = the engine default; passing `false` explicitly also counts.

### 3. Output (sink): a file, bytes, or an io.Writer

```go
datasource.FilePathSink("output.pdf") // to disk (parent dirs are created)
&out                                 // into memory; read it back with out.Bytes()
datasource.NewWriterSink(w)          // straight into any io.Writer (http.ResponseWriter, buffer, socket)
```

`BytesSink` and `WriterSink` are the two keys to "most flexible": they let the conversion result skip the disk entirely and flow straight into an HTTP response, object storage, a zip, and so on. All three are different implementations of the same `sink` parameter, so anything that accepts `[]byte` or an `io.Writer` can be plugged into converter.

---

## 4. Real-environment examples: one `Convert`, many shapes

### A. Convert in memory; the download endpoint streams PDF

The common server-side scenario — a user uploads xlsx, the endpoint returns PDF, no temp files:

```go
func handleExcelToPDF(w http.ResponseWriter, r *http.Request) {
	var out datasource.BytesSink
	if err := converter.Convert(
		datasource.NewReaderSource(r.Body),    // input: the request body
		pdf.New(pdf.WithOnePagePerSheet(true)), // format: PDF, one page per sheet
		&out,                                   // output: memory
	); err != nil {
		http.Error(w, err.Error(), http.StatusBadRequest)
		return
	}
	w.Header().Set("Content-Type", "application/pdf")
	w.Header().Set("Content-Disposition", `attachment; filename="report.pdf"`)
	_, _ = w.Write(out.Bytes())
}
```

To hand out CSV / Excel / HTML instead, swap only the `saveOption` (for example `formats.Get("csv")`); nothing else changes.

### B. Read once, batch-export to many formats

Read the input a single time, then reuse it through `BytesSource` for any number of outputs — ideal for "export every report":

```go
data, err := os.ReadFile("input.xlsx")
if err != nil {
	log.Fatal(err)
}

for _, ext := range []string{"pdf", "csv", "html", "json", "md"} {
	var out datasource.BytesSink
	if err := converter.Convert(datasource.BytesSource(data), formats.Get(ext), &out); err != nil {
		log.Fatalf("%s: %v", ext, err)
	}
	if err := os.WriteFile("output."+ext, out.Bytes(), 0o644); err != nil {
		log.Fatal(err)
	}
}
```

### C. "Save as" within the spreadsheet family (migration / compatibility)

Format migration inside the spreadsheet family, e.g. `xlsx → xls`, `xlsx → xlsb`, `xlsx → ods`, and the templates `xltx`/`xltm` — still one `Convert`, only the `saveOption` changes:

```go
src := datasource.FilePathSource("input.xlsx")
converter.Convert(src, formats.Get("xls"),  datasource.FilePathSink("legacy.xls")) // legacy Excel
converter.Convert(src, formats.Get("xlsb"), datasource.FilePathSink("binary.xlsb")) // binary Excel
converter.Convert(src, formats.Get("ods"),  datasource.FilePathSink("open.ods"))    // OpenDocument
converter.Convert(src, formats.Get("xltx"), datasource.FilePathSink("template.xltx")) // template
```

### D. Tune options per format, only when needed

Functional options appear when you need fine control — note every call below is real and correctly typed:

```go
// PDF: one page per worksheet + a compliance level
pdf.New(
	pdf.WithOnePagePerSheet(true),
	pdf.WithCompliance(asposecells.PdfCompliance_Pdf15),
	pdf.WithProducer("my-report-app"),
)

// HTML: single file, images inlined as base64, CSS prefix (avoids style clashes)
html.New(
	html.WithSaveAsSingleFile(true),
	html.WithExportImagesAsBase64(true),
	html.WithCellCssPrefix("tbl_"),
)

// CSV: custom separator and encoding
// (note: the separator is a byte, and the encoding is an EncodingType, not a string)
csv.New(
	csv.WithSeparator(';'),
	csv.WithEncoding(asposecells.EncodingType_UTF8),
)

// Image: pick the type (the only switch the image package currently exposes)
image.New(image.WithImageType("png")) // or "jpg" / "svg" / "gif" / "tif" / …
```

> Tip: for single-file outputs such as CSV/HTML/image, the engine exports only the worksheet it considers "active" by default. To control the scope, use the format's own options — e.g. `html.WithExportActiveWorksheetOnly(true)` (only the active sheet?), `html.WithShowAllSheets(true)`, or `csv.WithExportAllSheets(true)`. Pixel-level image quality, choosing exactly which sheets are rendered, and other deep controls still exist at the engine layer, but the current `image` package only exposes the `WithImageType` switch — reach for the underlying binding if you need more; the toolkit will keep folding options in over time.

### E. Batch conversion over a directory

The everyday "convert this folder of Excels to PDF" job becomes very short with extension inference:

```go
func convertAll(inputDir, outputDir string) error {
	if err := os.MkdirAll(outputDir, 0o755); err != nil {
		return err
	}
	return filepath.Walk(inputDir, func(path string, info os.FileInfo, err error) error {
		if err != nil || info.IsDir() {
			return err
		}
		ext := strings.ToLower(filepath.Ext(path))
		if ext != ".xlsx" && ext != ".xls" && ext != ".xlsm" && ext != ".csv" {
			return nil
		}
		out := filepath.Join(outputDir, strings.TrimSuffix(filepath.Base(path), ext)+".pdf")
		if err := convert(path, out); err != nil { // reuse the convert from section 2
			return fmt.Errorf("convert %s: %w", path, err)
		}
		fmt.Printf("done: %s -> %s\n", path, out)
		return nil
	})
}
```

### F. Classifying errors

converter errors are classifiable sentinel errors; pair them with `errors.Is` for precise handling:

```go
err := convert("in.xlsx", "out.xyz")
switch {
case errors.Is(err, toolkiterrors.ErrUnsupportedFormat):
	log.Println("this extension has no registered output format")
case errors.Is(err, os.ErrNotExist):
	log.Println("the input file does not exist")
default:
	log.Printf("conversion failed: %v", err)
}
```

`formats.List()` returns every currently registered extension (sorted, concurrency-safe) — handy for generating a "supported formats" list or an error hint.

---

## 5. Supported formats at a glance (32 registered by default)

| Family | Output extensions | Tuning package |
| --- | --- | --- |
| Excel | `xlsx` `xlsm` `xltx` `xltm` `xls` `xlsb` | `saveoptions/ooxml`, `xls`, `xlsb` |
| OpenDocument | `ods` `xml` `dif` `dbf` | `saveoptions/ods`, `xml`, `dif`, `dbf` |
| Text / data | `csv` `txt` `tsv` `json` `sql` | `saveoptions/csv`, `txt`, `json`, `sqlScript` |
| Documents / layout | `pdf` `docx` `pptx` `xps` `pcl` `epub` | `saveoptions/pdf`, `docx`, `pptx`, `xps`, `pcl`, `ebook` |
| Web | `html` `md` | `saveoptions/html`, `markdown` |
| Images | `png` `jpg` `jpeg` `gif` `bmp` `svg` `emf` `tif` `tiff` | `saveoptions/image` |

To keep the binary lean, import only the subpackages you use (such as `pdf`); for convenience or generic dispatch, blank-import `…/register`. Importing a subpackage triggers its registration; registry lookups are case-insensitive; an unregistered extension returns `nil` (not an error), leaving the fallback decision to the caller.

---

## 6. Best practices and caveats

1. **One entry point — do not reinvent it.** File-to-file uses `Convert(FilePathSource, formats.Get(ext)/constructor, FilePathSink)`; only switch sinks when you need bytes or a stream. Do not hand-roll `os.Open → ReadAll → … → os.WriteFile`; the source/sink manage resources and create parent directories for you.
2. **Concurrency and licensing.** The underlying engine is native: do **not** share one `SaveOption`/workbook object across goroutines in the same process; for bulk work process files sequentially or parallelize across processes. The evaluation build caps loads (about one hundred `NewWorkbook_Stream` loads per process) — for large benchmarks, restart the process to reset the counter.
3. **Read the input once, then reuse it.** To emit several formats from one input, `os.ReadFile` it into a `[]byte` first and convert repeatedly with `BytesSource`, avoiding repeated disk reads and reloads.
4. **Streams get buffered.** `NewReaderSource` buffers on first access, so it suits one-shot inputs like an HTTP request body, not true streaming of very large files (use `FilePathSource` pointing straight at disk for those).
5. **Know who decides the output layout.** `Convert` is "one input → one output file". If you need **one file per worksheet** or a **zip**, that is `manipulator.Split` with a `FolderSink`/`ZipSink` — do not force it through `Convert`.

## Reference

- Full API reference: [docs/converter.md](../docs/converter.md)
- The `formats` registry and custom formats: [docs/formats.md](../docs/formats.md), [docs/register.md](../docs/register.md)
- Runnable sample: [examples/convert/main.go](../examples/convert/main.go)
