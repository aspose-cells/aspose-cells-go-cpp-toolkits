package examples

import (
	"log"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/core"
)

// Cleanup releases the native engine. Examples defer it at the top of main so
// the process exits cleanly under engine v26.10.0, which crashes during exit
// when its process-wide state is left to the C runtime. Calling it more than
// once is safe.
func Cleanup() { core.Cleanup() }

// Fatal logs v and exits with status 1 after releasing the engine. Examples use
// it in place of log.Fatal: log.Fatal calls os.Exit directly, which skips
// deferred functions and therefore skips cleanup.
func Fatal(v ...any) {
	core.Cleanup()
	log.Fatal(v...)
}

// Fatalf is Fatal with printf-style formatting.
func Fatalf(format string, v ...any) {
	core.Cleanup()
	log.Fatalf(format, v...)
}
