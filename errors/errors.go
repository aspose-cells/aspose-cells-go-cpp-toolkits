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

	// ErrWorksheetNotFound is returned when a worksheet is referenced by name
	// but no worksheet with that name exists in the workbook.
	ErrWorksheetNotFound = errors.New("worksheet not found")

	// ErrDataSourceNil is returned when a nil datasource.DataSource is passed
	// to a toolkit entry point.
	ErrDataSourceNil = errors.New("data source is nil")

	// ErrDataSinkNil is returned when a nil datasource.DataSink is passed to a
	// toolkit entry point.
	ErrDataSinkNil = errors.New("data sink is nil")

	// ErrNoSources is returned when a merge entry point receives no input
	// sources, which would otherwise silently produce an empty workbook.
	ErrNoSources = errors.New("no data sources to merge")

	// ErrInvalidCellRef is returned when a cell reference string (e.g. "B3")
	// cannot be parsed into a cell coordinate.
	ErrInvalidCellRef = errors.New("invalid cell reference")

	// ErrInvalidRange is returned when a cell range has an invalid shape, e.g. a
	// range whose start cell lies below or to the right of its end cell, or an
	// area string with more than one ":" separator.
	ErrInvalidRange = errors.New("invalid cell range")

	// ErrLicenseInvalid is returned when the Aspose.Cells license cannot be
	// created or applied (e.g. the file is missing, unreadable, or invalid).
	ErrLicenseInvalid = errors.New("invalid license")

	// ErrColumnNotFound is returned when a struct field mapped by a row reader
	// (query.ReadRows) has no matching column in the worksheet's header row, or
	// when the header row is empty so no column mapping can be established.
	ErrColumnNotFound = errors.New("column not found in header row")
)
