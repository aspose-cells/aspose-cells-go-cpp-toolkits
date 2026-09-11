# core

Package core holds engine-level configuration.

Currently it exposes `SetLicense` for licensing the underlying Aspose.Cells engine before any spreadsheet is processed.

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
err := core.SetLicense("path/to/license.lic")
if err != nil {
    if errors.Is(err, core.ErrLicenseInvalid) {
        log.Println("Warning: Using evaluation version")
    } else {
        log.Fatalf("Failed to set license: %v", err)
    }
}
```

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