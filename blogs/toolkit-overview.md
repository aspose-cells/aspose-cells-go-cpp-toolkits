---
title: "Aspose.Cells for Go via C++ Toolkits: Excel, the Go way"
description: "An introduction to Aspose.Cells for Go via C++ Toolkits, a Go wrapper around the Aspose.Cells engine. Read, edit, convert, merge and split Excel (xlsx/xls/csv/pdf) with a small DataSource/DataSink API and a declarative DSL instead of the C++-style binding — with runnable examples."
keywords:
  - "Aspose.Cells for Go"
  - "Aspose.Cells for Go via C++ Toolkits"
  - "Excel Go library"
  - "Excel Go wrapper"
  - "go excel toolkit"
  - "Excel to PDF Go"
  - "Aspose.Cells Go binding"
  - "read write Excel Go"
---
# Aspose.Cells for Go via C++ Toolkits: Excel, the Go way

There is no shortage of Go libraries that touch Excel, but most are simply the official SDK transcribed en masse with `go:generate` — the calls still taste like C++: handles everywhere, `GetXxx`/`SetXxx`, and manual object-lifecycle management. **Aspose.Cells for Go via C++ Toolkits** does the opposite: it still runs on the Aspose.Cells engine, but folds the 80% of operations you perform every day into a **small set of Go-flavored composite functions**. This article walks you through its main capabilities with programs you can run as-is.

## 1. What it is built on

It is a wrapper toolkit on top of the official [Aspose.Cells for Go via C++](https://products.aspose.com/cells/go-cpp/) binding (`github.com/aspose-cells/aspose-cells-go-cpp`). The C++-style interface underneath is replaced:

| | Raw binding | Toolkits |
| --- | --- | --- |
| Open a file | `NewWorkbook_String`, then chain `GetWorksheets().Get_Int(0)` all the way down | one value: `datasource.FilePathSource("a.xlsx")` |
| Write one cell | `GetCells().Get_Int_Int(r,c).PutValue(...)`, grabbing objects layer by layer | `editor.SetCellValue(r, c, v)` — declarative |
| Save | one `Save_Xxx` per `SaveFormat` | pick a `saveOption` + sink for the target format |
| Read a whole sheet back | walk `GetRows()`/`GetCells()` by hand | `query.ReadRows[T]` reflects once into `[]struct` |
| Handle errors | every call returns `(obj, err)`; the semantics blur together | `errors.Is(err, toolkiterrors.ErrXxx)` classifies precisely |

The engine keeps all its power (read/write xlsx/xls/csv/pdf/…, formulas, charts, styles); it is just wrapped in a **composition layer**: `DataSource`/`DataSink` unify "file, memory, stream", and a handful of entry functions plus functional options replace hundreds of scattered methods. Requirements: **Go 1.21+, Windows x64 / Linux x64**.

## 2. What makes it distinctive

1. **I/O is fully abstracted — one signature goes anywhere.** Files, `[]byte`, and any `io.Reader`/`io.Writer` implement the same `DataSource`/`DataSink`, so the same function can write to disk or stream straight into an HTTP response.
2. **A declarative editing DSL.** `editor.InWorksheet(sheet, actions…)` states *what to do on which worksheet* instead of grabbing handles step by step.
3. **Typed read/write with reflection mapping.** `ReadRows[T]`/`WriteRows` bind struct fields to header columns via `excel:"column"` tags — a sheet becomes a `[]struct` and back again.
4. **Options, not overloads.** Want one page per worksheet in the PDF? `pdf.New(pdf.WithOnePagePerSheet(true))`. Zero configuration is the default use.
5. **Classifiable errors.** Shared sentinel errors combine with `errors.Is` for "unsupported format / worksheet not found / nil data source…".
6. **It does not mirror the underlying object model.** You program against intent — read a table, change a table, convert a format — not against Aspose's class tree.

## 3. Quick start

Create a project and pull the dependency:

```console
$ go mod init excel-demo
$ go get github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26@latest
```

Toolkits drive the engine through a C++ shared library, so put the binding's `lib` directory on `PATH` before running (`win_x86_64` on Windows, `linux_x64` on Linux; see the repo README):

```bash
# Windows (Git Bash)
export PATH="$HOME/go/pkg/mod/github.com/aspose-cells/aspose-cells-go-cpp/v26@v26.7.0/lib/win_x86_64:$PATH"
# Linux
export PATH="$HOME/go/pkg/mod/github.com/aspose-cells/aspose-cells-go-cpp/v26@v26.7.0/lib/linux_x64:$PATH"
```

> Without a License the engine runs in **evaluation mode**: output carries watermarks/warning sheets, there is a load-count budget, and cells/worksheet names are occasionally nudged (never in the saved bytes). The demos below tolerate that; for production configure `core.SetLicense` as described in the [README](../README.md).

### Hello, world: write it, read it back

Save the whole file as `main.go` and run `go run .` — it needs **no sample data file**, it builds a sheet in memory first:

```go
package main

import (
	"fmt"
	"log"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/query"
)

func main() {
	seed, err := datasource.NewEmptyWorkbook() // blank workbook: in memory, not counted as a load
	if err != nil {
		log.Fatal(err)
	}

	out, err := editor.EditSpreadsheet( // declarative: write two cells on the first sheet
		seed,
		editor.InWorksheet(0,
			editor.SetCellValue(0, 0, "Hello"),
			editor.SetCellValue(0, 1, "World"),
		),
	)
	if err != nil {
		log.Fatal(err)
	}

	// Read B1 back to verify. Evaluation mode (no License) can occasionally read
	// a cell as garbage, so retry up to 4 times; licensed runs get it first try.
	got := ""
	for i := 0; i < 4 && got != "World"; i++ {
		v, err := query.ReadCell(datasource.BytesSource(out), "B1", query.WithSheetIndex(0))
		if err == nil {
			got, _ = v.String()
		}
	}
	if got != "World" {
		log.Fatalf("evaluation-mode read-back was corrupted (got %q)", got)
	}
	fmt.Println("B1 = World") // expected output: B1 = World
}
```

Note there is no `os.Open`/`os.Create` and none of that `GetWorksheets().Get_Int(0).GetCells()…` object-chasing — input and output are `BytesSource`, writing is the DSL, and reading grabs a cell by reference ("B1"). Even "where does the empty workbook come from" is folded into `datasource.NewEmptyWorkbook()`: it creates a blank sheet in memory, so you never reach for the underlying `asposecells.NewWorkbook`.

## 4. All capabilities in one run

The longer program below strings together the toolkit's main traits and is **self-contained and directly runnable** too. It does five things: 1) build a styled table (merged cells and a comment), 2) read it back typed into `[]struct`, 3) convert it to PDF and CSV, 4) merge two workbooks and split them by worksheet, 5) export JSON and import CSV.

```go
package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/converter"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/manipulator"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/query"
	_ "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/register" // registers every output format
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/ooxml"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions/pdf"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/transfer"
)

// Row is the Go shape of one spreadsheet row; the excel tag binds each field
// to a header column (case-insensitive).
type Row struct {
	Item string `excel:"item"`
	Qty  int    `excel:"qty"`
}

func main() {
	out := "out"
	if err := os.MkdirAll(out, 0o755); err != nil {
		log.Fatal(err)
	}

	// A blank workbook is the seed for "start from nothing": it lives purely in
	// memory and can be reused as an input any number of times.
	seed, err := datasource.NewEmptyWorkbook()
	if err != nil {
		log.Fatal(err)
	}

	// ---- 1) build a table: declarative editing ----
	edited, err := editor.EditSpreadsheet(
		seed, // seed: one blank workbook
		editor.InWorksheet(0, // target the first worksheet
			editor.SetCellValue(0, 0, "Item"),
			editor.SetCellValue(0, 1, "Qty"),
			editor.SetCellValue(1, 0, "Apple"),
			editor.SetCellValue(1, 1, 12),
			editor.SetCellValue(2, 0, "Banana"),
			editor.SetCellValue(2, 1, 30),
			editor.SetStyle(0, 0, 1, 2, // header: bold, red
				editor.WithFontIsBold(true),
				editor.WithFontColor("red"),
			),
			editor.SetCellComment(1, 0, "best seller"),     // comment on A2
			editor.SetValue(0, 2, 1, 2, "prices in stock"), // C1:D1 filler
			editor.Merge(0, 2, 1, 2, true),                 // then merge C1:D1
		),
	)
	if err != nil {
		log.Fatal("edit: ", err)
	}
	fmt.Println("[1] edited: bold red header, A2 comment, C1:D1 merged")

	// ---- 2) read back: typed rows as []struct ----
	rows, err := readRowsStable(edited)
	if err != nil {
		log.Fatal("read rows: ", err)
	}
	fmt.Printf("[2] typed read-back, %d data rows:\n", len(rows))
	for _, r := range rows {
		fmt.Printf("    %-7s qty=%d\n", r.Item, r.Qty)
	}
	if s, ok := mustCell(edited, "C1"); ok {
		fmt.Printf("    merged C1 = %s\n", s)
	}
	if note, err := readCommentStable(edited, "A2"); err != nil {
		log.Print("comment read: ", err)
	} else {
		fmt.Printf("    A2 comment = %q\n", note)
	}

	// ---- 3) convert: the same in-memory bytes -> PDF / CSV ----
	var pdfOut datasource.BytesSink
	if err := converter.Convert(
		datasource.BytesSource(edited),
		pdf.New(pdf.WithOnePagePerSheet(true)),
		&pdfOut,
	); err != nil {
		log.Fatal("to pdf: ", err)
	}
	mustWrite(filepath.Join(out, "table.pdf"), pdfOut.Bytes())
	if err := converter.Convert(
		datasource.BytesSource(edited),
		formats.Get("csv"), // format looked up by name; csv.New() works too
		datasource.FilePathSink(filepath.Join(out, "table.csv")),
	); err != nil {
		log.Fatal("to csv: ", err)
	}
	fmt.Println("[3] converted: table.pdf / table.csv written")

	// ---- 4) merge, then split by worksheet ----
	two, err := editor.EditSpreadsheet(
		seed, // reuse the same blank seed to build a second sheet
		editor.InWorksheet(0,
			editor.SetCellValue(0, 0, "Region"),
			editor.SetCellValue(0, 1, "Target"),
			editor.SetCellValue(1, 0, "East"),
			editor.SetCellValue(1, 1, 90),
			editor.SetValue(3, 0, 1, 2, "note"), // merged note row below the table
			editor.Merge(3, 0, 1, 2, true),
		),
	)
	if err != nil {
		log.Fatal("edit second: ", err)
	}
	var merged datasource.BytesSink
	if err := manipulator.Merge(
		[]datasource.DataSource{datasource.BytesSource(edited), datasource.BytesSource(two)},
		ooxml.New(), &merged,
	); err != nil {
		log.Fatal("merge: ", err)
	}
	mustWrite(filepath.Join(out, "merged.xlsx"), merged.Bytes())
	if err := manipulator.Split(
		datasource.BytesSource(merged.Bytes()),
		ooxml.New(),
		datasource.FolderSink(filepath.Join(out, "split")), // one file per worksheet
	); err != nil {
		log.Fatal("split: ", err)
	}
	fmt.Println("[4] merged -> merged.xlsx; split by worksheet into out/split/")

	// ---- 5) data exchange: export a range to JSON + import CSV ----
	var jsonOut datasource.BytesSink
	if err := transfer.ExportRangeToJson(
		datasource.BytesSource(edited), &jsonOut,
		transfer.WithSheetIndex(0), transfer.WithStartCell("A1"), transfer.WithEndCell("B3"),
	); err != nil {
		log.Fatal("range to json: ", err)
	}
	mustWrite(filepath.Join(out, "range.json"), jsonOut.Bytes())

	csvData := []byte("City,Pop\nTokyo,37\nOsaka,19\n")
	if err := transfer.ImportCSV(
		seed, // import into a blank workbook
		datasource.BytesSource(csvData),
		datasource.FilePathSink(filepath.Join(out, "imported.xlsx")),
		transfer.WithSheetIndex(0), transfer.WithBeginCell(0, 0),
		transfer.WithConvertNumeric(true), transfer.WithSeparator(","),
	); err != nil {
		log.Fatal("import csv: ", err)
	}
	fmt.Println("[5] range.json exported; CSV imported -> imported.xlsx")
	fmt.Println("all done ✔")
}

// ---- small helpers used above ----

func mustWrite(path string, b []byte) {
	if err := os.WriteFile(path, b, 0o644); err != nil {
		log.Fatal(err)
	}
}

// The three "Stable" helpers below exist only for evaluation-mode tolerance:
// without a License the engine occasionally reads a cell value as garbage
// (about 1 in 3 loads on tiny files), and if it lands on a header cell,
// ReadRows would report "column not found". Licensed users can call
// query.ReadRows / query.ReadCell / query.ReadCellComment directly.
func readRowsStable(b []byte) ([]Row, error) {
	var rows []Row
	var err error
	for i := 0; i < 3; i++ {
		rows, err = query.ReadRows[Row](datasource.BytesSource(b), query.WithSheetIndex(0))
		if err == nil {
			return rows, nil
		}
	}
	return rows, err
}

func mustCell(b []byte, ref string) (string, bool) {
	for i := 0; i < 3; i++ {
		v, err := query.ReadCell(datasource.BytesSource(b), ref, query.WithSheetIndex(0))
		if err == nil {
			return v.String()
		}
	}
	return "", false
}

func readCommentStable(b []byte, ref string) (string, error) {
	var out string
	var err error
	for i := 0; i < 3; i++ {
		out, err = query.ReadCellComment(datasource.BytesSource(b), ref, query.WithSheetIndex(0))
		if err == nil {
			return out, nil
		}
	}
	return out, err
}
```

Running it prints the following (excerpt) and, under `out/`, produces `table.pdf`, `table.csv`, `merged.xlsx`, `out/split/*.xlsx`, `range.json`, and `imported.xlsx`:

```console
[1] edited: bold red header, A2 comment, C1:D1 merged
[2] typed read-back, 2 data rows:
    Apple   qty=12
    Banana  qty=30
    merged C1 = prices in stock
    A2 comment = "best seller"
[3] converted: table.pdf / table.csv written
[4] merged -> merged.xlsx; split by worksheet into out/split/
[5] range.json exported; CSV imported -> imported.xlsx
all done ✔
```

For comparison: that single program already uses entry points from six packages, and they all share the same `DataSource`/`DataSink` intuition — an input is always "something that can be opened", an output is always "something that can catch bytes".

> Tip: the final step `Split` breaks the merged result into `out/split/`, one file per worksheet. In **unlicensed (evaluation) mode** the engine injects a worksheet named `Evaluation Warning…` into the merged output, so the split folder contains extra `Evaluation Warning.xlsx` files; with a License you only see real worksheet names such as `Sheet1.xlsx`.

## 5. Capability map (few entry points, plain semantics)

| What you want | Package | One representative use |
| --- | --- | --- |
| Change a sheet (values/styles/merge/rows/cols/worksheets/comments/named ranges/encryption) | `editor` | `editor.EditSpreadsheet(src, editor.InWorksheet(0, editor.SetCellValue(0,0,"Hi")))` |
| Read a sheet (cell/range/whole sheet/merged areas/typed rows/named ranges/comments) | `query` | `query.ReadRows[T](src, query.WithSheetIndex(0))` |
| Convert format (xlsx/csv/pdf/html/… 32 outputs) | `converter` + `saveoptions/*` | `converter.Convert(src, pdf.New(pdf.WithOnePagePerSheet(true)), &sink)` |
| Merge several workbooks / split by worksheet | `manipulator` | `manipulator.Merge(sources, opt, sink)` |
| Import/export CSV · JSON · XML | `transfer` | `transfer.ExportRangeToJson(src, &sink, transfer.WithStartCell("A1"))` |
| Resolve an output format by extension / register custom formats | `formats` | `opt := formats.Get("xlsx")` |
| Create data sources and sinks | `datasource` | `FilePathSource` / `BytesSource` / `NewReaderSource` / **`NewEmptyWorkbook`** (start from a blank book) / `FilePathSink` / `BytesSink` / `NewWriterSink` |

## 6. What else the repo ships

Besides the library itself, the repo carries a **set of self-contained, directly runnable** examples (`examples/convert`, `edit`, `query`, `merge-split`, `transfer`), each reading the repo's bundled sample data and writing its own `out/`:

```bash
./examples/run.sh            # run every example in order
./examples/run.sh query      # or just one
```
