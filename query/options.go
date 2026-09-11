package query

import (
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
)

// Option configures a query entry point.
type Option func(*options)

type options struct {
	sheet     cells.Sheet
	trimSpace bool
}

func defaultOptions() *options {
	// Default to the first worksheet by index. The default first sheet's name
	// is the name evaluation mode most often corrupts at load time, so a
	// name-based default is deliberately avoided.
	return &options{sheet: cells.FirstSheet}
}

// WithSheet selects the worksheet to read from by name.
//
// Note: in evaluation mode the engine occasionally corrupts a worksheet's name
// when the workbook is loaded (observed ~2% of loads, any sheet, not just the
// default first sheet), so a name-based lookup may fail with
// ErrWorksheetNotFound even for a sheet that exists. Prefer WithSheetIndex,
// which is immune to name corruption.
func WithSheet(name string) Option {
	return func(o *options) { o.sheet = cells.Sheet{Name: name} }
}

// WithSheetIndex selects the worksheet to read from by its zero-based index.
// Index-based lookup is immune to the evaluation-mode load-time name
// corruption, so it is the recommended way to target a sheet.
func WithSheetIndex(i int) Option {
	return func(o *options) { o.sheet = cells.Sheet{UseIndex: true, Index: i} }
}

// WithTrimSpace controls whether leading and trailing whitespace on string
// cells is trimmed before the value is returned. Defaults to false.
func WithTrimSpace(b bool) Option {
	return func(o *options) { o.trimSpace = b }
}

func applyOptions(o *options, opts []Option) {
	for _, opt := range opts {
		if opt != nil {
			opt(o)
		}
	}
}
