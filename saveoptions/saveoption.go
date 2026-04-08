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
// without modifying core conversion functions like ConvertSpreadsheet or ConvertToWriter.
type SaveOption interface {
	// Apply transforms the given input byte slice (representing a spreadsheet) into the desired output format.
	// It returns the resulting byte slice and any error encountered during processing.
	Apply([]byte) ([]byte, error)
	GetFormat() string
}
