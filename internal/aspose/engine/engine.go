package engine

import (
	"reflect"
	"runtime"

	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// The native engine is not safe to enter from more than one goroutine at a
// time. Two concurrent loads wedge it outright (measured: four goroutines
// loading in a loop hang on every run), and a load that overlaps a finalizer
// freeing an object dies with an access violation inside the load path
// (measured: 4 of 15 full test runs). Every public entry point that can reach
// the engine holds this lock for the duration of its work; see lock.go for why
// it is re-entrant.
//
// The binding's own finalizers are the other way into the engine, and they do
// not take this lock. That is why the constructors below disarm them.

// WithEngine runs fn with exclusive access to the native engine.
func WithEngine(fn func() error) error {
	LockEngine()
	defer UnlockEngine()
	return fn()
}

// LockEngine takes exclusive access to the native engine. The lock is
// re-entrant per goroutine, so a public entry point may take it even when the
// call it makes into another package will take it again. Every call must be
// paired with UnlockEngine on the same goroutine.
func LockEngine() { engineLock.lock() }

// UnlockEngine releases one level of the lock taken by LockEngine.
func UnlockEngine() { engineLock.unlock() }

// disarm stops the finalizer the binding installed on obj from ever running.
//
// The binding sets runtime.SetFinalizer on every object it returns, and those
// finalizers call straight into the engine from Go's finalizer goroutine,
// which does not take engineMu. Left armed, they are a second goroutine
// entering the engine at arbitrary points — the race this package exists to
// close. Handles derived from a workbook (worksheets, cells, styles, charts)
// are disarmed but never freed by the toolkit: the engine owns them, and
// freeing them through Delete_* is what turns a stale wrapper into a
// use-after-free of engine memory.
func disarm(obj any) {
	runtime.SetFinalizer(obj, nil)
}

// Disarm stops the binding's finalizer from running for an engine handle that
// the toolkit does not own. See disarm for why this is necessary.
func Disarm(obj any) { disarm(obj) }

// Derive wraps a binding call that returns a handle the engine owns, so the
// handle arrives already disarmed: Derive(ws.GetCells()) reads as one call and
// spares every accessor its own error-check-then-disarm boilerplate.
//
// This is not an optimisation, it is the fix for the crashes locking could not
// reach. The binding's finalizers are destructive, not merely wasteful:
// Delete_Worksheet really does destroy the worksheet, Delete_Cell the cell. A
// wrapper dropped inside a public entry point is unreachable, so the next
// collection destroys an object the engine still has a pointer to, and the call
// after that walks a freed object. Measured with a probe that creates and drops
// handles in a loop: Delete_Worksheet, Delete_Cell and Delete_Charts each fault
// the engine within 20000 iterations. Disarming the same handles takes the
// probe from 1 pass in 5 to 5 in 5.
//
// Only engine-owned handles belong here. An option object the caller created
// (a save option, a load option) is owned by that caller and freed by its own
// finalizer, so disarming one would leak it instead of fixing anything.
func Derive[T any](obj T, err error) (T, error) {
	if err == nil {
		disarmValue(obj)
	}
	return obj, err
}

// disarmValue disarms obj when it is a pointer, and does nothing otherwise.
// Getter calls that read a handle and getter calls that read a scalar have the
// same shape at the call site, so the wrapper has to accept both: a range read
// through a getter is a plain value and has no finalizer to clear.
func disarmValue(obj any) {
	rv := reflect.ValueOf(obj)
	if rv.Kind() == reflect.Ptr && !rv.IsNil() {
		disarm(obj)
	}
}

// OpenWorkbook loads data into a workbook the toolkit owns, with the binding's
// finalizer disarmed so release happens only where the toolkit chooses. Release
// it with CloseWorkbook; a workbook returned by this function is never freed by
// the garbage collector.
func OpenWorkbook(data []byte) (*asposecells.Workbook, error) {
	wb, err := asposecells.NewWorkbook_Stream(data)
	if err != nil {
		return nil, err
	}
	disarm(wb)
	return wb, nil
}

// NewWorkbook creates a blank workbook the toolkit owns. Release it with
// CloseWorkbook.
func NewWorkbook() (*asposecells.Workbook, error) {
	wb, err := asposecells.NewWorkbook()
	if err != nil {
		return nil, err
	}
	disarm(wb)
	return wb, nil
}

// OpenWorkbookFile loads a workbook from a path on disk, with the binding's
// finalizer disarmed. Release it with CloseWorkbook.
func OpenWorkbookFile(path string) (*asposecells.Workbook, error) {
	wb, err := asposecells.NewWorkbook_String(path)
	if err != nil {
		return nil, err
	}
	disarm(wb)
	return wb, nil
}

// OpenWorkbookWithOptions is OpenWorkbook for a load that needs engine load
// options. The caller keeps ownership of lo.
func OpenWorkbookWithOptions(data []byte, lo *asposecells.LoadOptions) (*asposecells.Workbook, error) {
	wb, err := asposecells.NewWorkbook_Stream_LoadOptions(data, lo)
	if err != nil {
		return nil, err
	}
	disarm(wb)
	return wb, nil
}

// CloseWorkbook releases a workbook opened through this package. DeleteWorkbook
// disarms the binding's finalizer and frees the native object in one step, so
// this is safe to call on any workbook OpenWorkbook returned.
//
// Call it exactly once per workbook. DeleteWorkbook nils the handle's pointer
// after freeing, so a second call arrives at the engine's allocator with a null
// pointer; the damage is silent and shows up much later as an access violation
// inside an unrelated call (measured: 2 of 30 full test runs crashed in a Save,
// two test files downstream of the double release). Losing a workbook forever
// costs only memory, so when in doubt leak rather than guess at a second
// release.
func CloseWorkbook(wb *asposecells.Workbook) {
	if wb == nil {
		return
	}
	asposecells.DeleteWorkbook(wb)
}
