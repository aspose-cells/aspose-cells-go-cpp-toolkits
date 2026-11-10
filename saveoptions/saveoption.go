// Package saveoptions defines the SaveOption contract and hosts the
// per-format option packages.
//
// A SaveOption turns raw spreadsheet bytes into a specific output format via
// Apply, and reports the target format through GetFormat. Each subpackage under
// saveoptions implements SaveOption for one family of formats and exposes
// functional options (With*) to configure the export.
package saveoptions

// SaveOption defines the behavior for transforming spreadsheet data into a specific output format.
// Implementations of this interface encapsulate format-specific conversion logic (e.g., PDF, XLSX, CSV).
//
// The Apply method takes raw input data (typically the bytes of a source spreadsheet) and returns
// the converted output data in the target format. It may also return an error if the transformation fails.
//
// Example implementations might include:
//   - PDFSaveOption: converts to PDF
//   - CSVSaveOption: exports to CSV
//   - XLSXSaveOption: re-saves as XLSX with specific settings
//
// This interface enables a flexible, decoupled design where conversion logic can be extended
// without modifying core conversion functions like converter.Convert, manipulator.Merge, or
// manipulator.Split.
type SaveOption interface {
	// Apply transforms the given input byte slice (representing a spreadsheet) into the desired output format.
	// It returns the resulting byte slice and any error encountered during processing.
	Apply([]byte) ([]byte, error)
	GetFormat() string
}

// WorkbookApplier is an optional extension of SaveOption for implementations
// that can save a workbook the caller has already loaded.
//
// Converter, merge, and split load a workbook to do their work and then need it
// written in a target format. With only Apply they have to serialize that
// workbook to bytes and pass the bytes to Apply, which reloads a second
// workbook just to save it. Implementing WorkbookApplier lets those entry
// points skip the serialize/reload cycle. It also lets them honor load options
// (a source password) on a source that Apply could not open by itself.
//
// Every SaveOption shipped by this module implements WorkbookApplier. Apply
// remains the interface callers program against, so a third-party
// implementation that only provides Apply keeps working, at the cost of the
// extra serialization.
type WorkbookApplier interface {
	SaveOption
	// SaveWorkbook writes workbook in this option's target format and returns
	// the bytes.
	//
	// workbook is an *asposecells.Workbook; the type is any so that the
	// toolkit's public API does not name a native binding type. An
	// implementation that is handed anything else returns an error wrapping
	// ErrInvalidWorkbook. The caller keeps ownership of workbook and closes it
	// after SaveWorkbook returns; implementations must not free it.
	//
	// Implementations take the toolkit's re-entrant engine lock themselves, so
	// SaveWorkbook may be called while the caller already holds it.
	SaveWorkbook(workbook any) ([]byte, error)
}
