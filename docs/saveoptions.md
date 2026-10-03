# saveoptions

## Types

### SaveOption

```go
type SaveOption interface {
	// Apply transforms the given input byte slice (representing a spreadsheet) into the desired output format.
	// It returns the resulting byte slice and any error encountered during processing.
	Apply([]byte) ([]byte, error)
	GetFormat() string
}
```

SaveOption defines the behavior for transforming spreadsheet data into a specific output format. Implementations of this interface encapsulate format-specific conversion logic (e.g., PDF, XLSX, CSV).

The Apply method takes raw input data (typically the bytes of a source spreadsheet) and returns the converted output data in the target format. It may also return an error if the transformation fails.

Example implementations might include:

  - PDFSaveOption: converts to PDF
  - CSVSaveOption: exports to CSV
  - XLSXSaveOption: re-saves as XLSX with specific settings

This interface enables a flexible, decoupled design where conversion logic can be extended without modifying core conversion functions like converter.Convert, manipulator.Merge, or manipulator.Split.


## Enum option names

Options that select one of the engine's enumerated values take the name as a
`string` — never the binding's enum type — so configuring a save does not require
importing the engine package. The vocabulary is the engine's own, spelled in
lower camel case after the constant suffix:

| Engine constant | Name |
| --- | --- |
| `asposecells.EncodingType_UTF8` | `"utf8"` |
| `asposecells.CellValueFormatStrategy_DisplayString` | `"displayString"` |
| `asposecells.EmfRenderSetting_EmfPlusPrefer` | `"emfPlusPrefer"` |
| `asposecells.PdfCompliance_PdfA1b` | `"pdfA1b"` |

Names are matched **case- and punctuation-insensitively**, the same rule `editor`
already applies to its own string enums, so `"UTF-8"`, `"utf8"` and `"utf-8"` are
one name. Each option's doc comment lists the names it accepts.

```go
csv.New(
    csv.WithEncoding("UTF-8"),          // or "ascii", "default", "unicode"
    csv.WithQuoteType("minimum"),       // or "always", "never", "normal"
    csv.WithFormatStrategy("displayString"),
)
```

An unrecognized name is an error (`errors.ErrInvalidEnumValue`), never a fallback
to the engine's default. The engine enumerates its own members, so a name it does
not recognize is a request for something that is not there, and quietly
substituting another value is how a mistyped option becomes a wrong document. The
check runs before the workbook is opened, so a bad name produces no output at all.

`saveoptions/image` is the one package whose format names are not member names:
the engine calls the JPEG member `Jpeg` and the TIFF member `Tiff`, but `"jpg"`
and `"tif"` name the same two formats and are what callers type, so both spellings
are accepted.

## Color option values

Options that take a color also take it in a form the toolkit resolves, never the
binding's `*asposecells.Color`:

| Form | Example |
| --- | --- |
| Go `color.Color` | `color.RGBA{R: 0x33, G: 0x66, B: 0x99, A: 0xFF}` |
| hex string | `"#336699"`, `"#33669980"` (8-digit is `RRGGBBAA`), with or without the `#` |
| color name | `"red"`, `"Light Sea Green"` |
| ARGB `int` | `0xFF336699` |

```go
pdf.New(pdf.WithGridlineColor("#336699"))
pdf.New(pdf.WithGridlineColor(color.NRGBA{R: 0x33, G: 0x66, B: 0x99, A: 0x80}))
```

`WithGridlineColor` in `docx`, `pcl`, `pdf`, `pptx` and `xps` takes any of these,
as do the color arguments of `editor` — one color vocabulary for the whole
toolkit. A value that is none of them is `errors.ErrInvalidColor`, reported
before any output is produced.

Note the eight-digit order: it is `RRGGBBAA`, the CSS order, which is what the
toolkit documents and what `editor` has always promised. The engine's own
`Color_FromHex` reads eight digits as `AARRGGBB`, so the toolkit reorders before
handing it over.

## Area option values

An option that limits work to a region takes an Excel-style area as a `string`,
not the binding's `CellArea`:

```go
json.New(json.WithExportArea("A1:C3"))
json.New(json.WithExportArea("B2"))     // a single cell is a one-cell area
```

The area must lie inside the worksheet grid. An area outside it, one spelled
backwards (`"C3:A1"`), or one that is not an area at all is reported as
`errors.ErrInvalidRange` / `errors.ErrInvalidCellRef` before any output is
produced. The engine does not check — it accepts an off-grid area and then
exports nothing — so this is the difference between an error and an empty file.

## Options awaiting a toolkit-native value

Eight further option families take a binding type that a string cannot express
(a `SheetSet`, a `PdfSecurityOptions`, a `RenderingWatermark`, a callback, …).
Their `With*` functions are **commented out** in the source, with a note on each
saying why and what to replace the parameter with; the `Config` field and the
`Apply` wiring are kept so restoring one is a matter of uncommenting and
changing the parameter type. This is deliberate rather than a gap in the build:
a caller who set one of those options would have to import the engine package,
which is exactly what the toolkit's public API must not require.
