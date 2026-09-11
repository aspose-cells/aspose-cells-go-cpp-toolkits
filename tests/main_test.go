package tests

import (
	"fmt"
	"os"
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/core"
)

// TestMain applies the Aspose.Cells license from the LicenseFilePath
// environment variable (falling back to the legacy LicensePath) before any
// test runs. Without a license the engine operates in evaluation mode, which
// caps the number of workbook loads per process and can corrupt worksheet
// names/values at load time (~2%), so setting LicenseFilePath to a valid .lic
// file makes the whole suite deterministic and quota-free. When neither
// variable is set the suite still runs in evaluation mode with its retry
// guards. A configured-but-invalid license path fails the run immediately:
// silently degrading to evaluation mode would produce confusing, unrelated
// quota failures later in the suite.
func TestMain(m *testing.M) {
	path := os.Getenv("LicenseFilePath")
	if path == "" {
		path = os.Getenv("LicensePath")
	}
	if path != "" {
		if err := core.SetLicense(path); err != nil {
			fmt.Fprintf(os.Stderr, "tests: applying license %q failed: %v\n", path, err)
			os.Exit(1)
		}
		fmt.Fprintf(os.Stderr, "tests: license applied from %q\n", path)
	} else {
		fmt.Fprintln(os.Stderr, "tests: no license configured (LicenseFilePath unset), running in evaluation mode")
	}
	os.Exit(m.Run())
}
