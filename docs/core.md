# core

Package core holds engine-level configuration.

It exposes `SetLicense` for licensing the underlying Aspose.Cells engine before any spreadsheet is processed, and `Cleanup` to release the engine's process-wide state before the program exits.

## Functions

### SetLicense

```go
func SetLicense(licensePath string) error
```

Loads and applies a license for the Aspose.Cells engine from the specified file path.

#### Parameters

- **licensePath**: The absolute or relative path to a valid Aspose.Cells license file (typically with a `.lic` extension). If empty, the `LicenseFilePath` environment variable is used (falling back to the legacy `LicensePath`). A missing or unreadable file is reported immediately as an error wrapping `ErrLicenseInvalid`; a readable file that contains an invalid license leaves subsequent operations in evaluation mode (e.g., with watermarks or feature limitations).

#### Returns

- **error**: `nil` when the license was applied; otherwise an error wrapping `ErrLicenseInvalid` (plus the underlying engine error). Callers that need to detect license failure should check the return value; a statement call without one keeps compiling and simply ignores the failure.

#### Notes

- This function configures the global license state for the underlying Aspose.Cells for C++ library. It should be called once at application startup before performing any conversion or manipulation operations.
- Calling `SetLicense` multiple times has no effect after a valid license is successfully loaded.
- If no license is set, the library operates in trial mode, which may impose restrictions such as watermarks on output documents or a limited worksheet size.

#### Example

```go
if err := core.SetLicense(os.Getenv("LicenseFilePath")); err != nil {
    log.Fatal(err)
}
```

#### License verification

To verify that the license was applied successfully, check the return value:

```go
import (
    "errors"
    "log"

    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/core"
    toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
)

err := core.SetLicense("path/to/license.lic")
if err != nil {
    if errors.Is(err, toolkiterrors.ErrLicenseInvalid) {
        log.Println("Warning: Using evaluation version")
    } else {
        log.Fatalf("Failed to set license: %v", err)
    }
}
```

`SetLicense` is `core`'s only exported symbol; the sentinel it wraps lives in the `errors` package, which the docs elsewhere import as `toolkiterrors`.

## Evaluation vs Licensed mode

### Evaluation mode

- Watermarks are inserted on saved files
- Limited to 100 file loads per process
- All functionality is available (no feature restrictions)

### Licensed mode

- No watermarks
- No file load limits
- Full production capabilities

## Best practices

1. **Call SetLicense once at startup**: License configuration is global. Call `SetLicense` once when your application starts.

2. **Always check the error**: Even though old code may ignore the return value, check the error to detect license failures:

   ```go
   // Correct
   if err := core.SetLicense(path); err != nil {
       log.Fatal(err)
   }
   
   // Avoid (legacy style)
   core.SetLicense(path) // errors are silently ignored
   ```

3. **Use environment variable for production**: Store the license file path in an environment variable rather than hardcoding it:

   ```go
   core.SetLicense(os.Getenv("ASPOSE_LICENSE_PATH"))
   ```

4. **Handle missing license gracefully**: For applications that support both evaluation and licensed modes:

   ```go
   if path := os.Getenv("LicenseFilePath"); path != "" {
       if err := core.SetLicense(path); err != nil {
           log.Printf("License error: %v; running in evaluation mode", err)
       }
   }
   ```

## Release the engine before exit

Starting with Aspose.Cells for Go via C++ **v26.10.0**, the native engine must
be shut down explicitly before the process exits. If it is not, the C runtime
tears down the engine's process-wide state while the Go runtime is exiting and
the program dies with an access violation (`Exception 0xc0000005`, `PC=0x0`)
after `main` has returned.

The toolkit exposes the binding's `Cleanup` as:

```go
func core.Cleanup()
```

Call it once, after the last toolkit operation and before the process exits.

### The recommended pattern

`os.Exit` and `log.Fatal` terminate the process immediately and do **not** run
deferred functions. Put `defer core.Cleanup()` in a function that returns to
`main`, and let `main` do the exiting:

```go
func main() {
    os.Exit(run())
}

func run() int {
    defer core.Cleanup() // runs when run returns, before os.Exit

    if err := core.SetLicense(os.Getenv("LicenseFilePath")); err != nil {
        log.Print(err)
        return 1
    }
    if err := work(); err != nil {
        log.Print(err)
        return 1
    }
    return 0
}
```

Returning an exit code instead of calling `log.Fatal` inside `run` is what
makes the deferred cleanup run. If the program installs its own signal
handler, call `core.Cleanup()` from that handler before exiting as well.

### Semantics

- **Terminal.** Every toolkit call after `Cleanup` returns is undefined.
  Finish all work and stop goroutines that still use the toolkit first.
- **Waits for in-flight work.** `Cleanup` takes the same engine lock as every
  other entry point, so it does not overlap a load, save, or handle release.
- **Idempotent.** Calling it more than once is safe, including when no workbook
  was ever opened.
- **Not required to close workbooks.** Close or finish workbooks through their
  normal entry points; `Cleanup` releases process-wide engine state, not
  individual documents.
