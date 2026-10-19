package core

import (
	"sync"

	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/engine"
)

// cleanupOnce makes the Go side idempotent: the binding's Cleanup is safe to
// call repeatedly, but there is no reason to re-enter the engine after its
// process-wide state has been released.
var cleanupOnce sync.Once

// Cleanup releases the native engine's process-wide state.
//
// Call it once, after the last toolkit operation has finished and before the
// process exits:
//
//	func main() {
//		os.Exit(run())
//	}
//
//	func run() int {
//		defer core.Cleanup()
//		if err := work(); err != nil {
//			log.Print(err)
//			return 1
//		}
//		return 0
//	}
//
// From Aspose.Cells for Go via C++ v26.10.0 the engine must be shut down
// explicitly. Skipping Cleanup leaves the teardown to the C runtime during
// process exit, which crashes the program with an access violation after main
// returns (observed on Windows as "Exception 0xc0000005" with PC=0x0 inside
// runtime.exit).
//
// Cleanup waits for in-flight toolkit operations to finish before releasing
// the engine, so it is safe to call while other work is still winding down.
// It is terminal: every toolkit call after it returns is undefined. Stop any
// goroutines that still use the toolkit before calling it, and do not call it
// from inside an operation.
//
// Cleanup is safe to call more than once, including when no workbook was ever
// opened.
func Cleanup() {
	cleanupOnce.Do(func() {
		// Take the engine lock for the same reason every entry point does:
		// Cleanup must not overlap a load, save, or handle release. The lock
		// is re-entrant per goroutine, so calling Cleanup from a deferred
		// function on the goroutine that ran the work is safe.
		_ = engine.WithEngine(func() error {
			asposecells.Cleanup()
			return nil
		})
	})
}
