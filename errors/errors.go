// Package errors defines sentinel errors returned by the toolkit.
//
// Returned errors are wrapped with %w so callers can classify failures with
// errors.Is / errors.As instead of matching on error message strings. For
// example, a failed conversion can be distinguished from an unsupported
// output format:
//
//	_, err := converter.ConvertSpreadsheetToFile("in.xlsx", "out.xyz")
//	if errors.Is(err, toolkiterrors.ErrUnsupportedFormat) {
//		// fall back to a different extension
//	}
//	if errors.Is(err, fs.ErrNotExist) {
//		// input file is missing (os errors propagate unchanged)
//	}
package errors

import "errors"

var (
	// ErrSaveOptionNil is returned when a nil saveoptions.SaveOption is passed
	// to a conversion or merge entry point.
	ErrSaveOptionNil = errors.New("save option is nil")

	// ErrUnsupportedFormat is returned when no output format is registered for
	// the requested file extension. The offending extension is included in the
	// wrapped message.
	ErrUnsupportedFormat = errors.New("unsupported output format")

	// ErrInvalidOutputPath is returned when an output path cannot be used to
	// infer a format, e.g. it has no file extension.
	ErrInvalidOutputPath = errors.New("invalid output path")

	// ErrInvalidSheetID is returned when a worksheet identifier is neither a
	// string (sheet name) nor an integer (sheet index).
	ErrInvalidSheetID = errors.New("invalid sheet identifier")

	// ErrInvalidValue is returned when a value of an unsupported type is
	// written to a cell or range.
	ErrInvalidValue = errors.New("invalid value")

	// ErrInvalidColor is returned when a color is given in an unsupported type.
	ErrInvalidColor = errors.New("invalid color value")

	// ErrInputIsFolder is returned when a file operation receives a directory
	// path where a file is required.
	ErrInputIsFolder = errors.New("input path is a folder, expected a file")
)
