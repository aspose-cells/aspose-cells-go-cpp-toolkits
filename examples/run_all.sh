#!/usr/bin/env bash
# run_all.sh - Build, vet, format-check, test, and run all examples

set -euo pipefail

script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$script_dir/.."

echo "=== 1. Build project ==="
go build ./...

echo "=== 2. Vet code ==="
go vet ./...

echo "=== 3. Format check ==="
# Check only packages in this module; ignored/generated trees such as
# agent/ResultOutput are outside the module and would fail gofmt -l .
package_dirs=$(go list -f '{{.Dir}}' ./...)
# List each package's .go files explicitly: gofmt recurses when handed a
# directory, which would pull in trees outside the module.
go_files=$(find $package_dirs -maxdepth 1 -type f -name '*.go')
unformatted=$(gofmt -l $go_files)
if [[ -n "$unformatted" ]]; then
    echo "Unformatted files:"
    echo "$unformatted"
    exit 1
fi
echo "Format check passed"

echo "=== 4. Resolve native library directory ==="
mod_dir=$(go list -m -f '{{.Dir}}' github.com/aspose-cells/aspose-cells-go-cpp/v26)
if [[ "$OSTYPE" == "msys" ]] || [[ "$(uname)" == "MINGW"* ]] || [[ "$(uname)" == "CYGWIN"* ]]; then
    # Go returns a native path (C:\...). Convert it to a POSIX path before
    # appending to PATH, otherwise the drive-letter colon is read as a list
    # separator and the DLL directory is never found.
    if command -v cygpath >/dev/null 2>&1; then
        mod_dir=$(cygpath -u "$mod_dir")
    fi
    lib_dir="$mod_dir/lib/win_x86_64"
    export PATH="$PATH:$lib_dir"
else
    lib_dir="$mod_dir/lib/linux_x86_64"
    export LD_LIBRARY_PATH="$lib_dir:$LD_LIBRARY_PATH"
fi
echo "Native library directory: $lib_dir"

echo "=== 5. Run tests ==="
go test -v ./...

echo "=== 6. Run examples ==="
echo "Running chart example..."
"$script_dir/run.sh" chart

echo "Running chart-export example..."
"$script_dir/run.sh" chart-export

echo "Running convert example..."
"$script_dir/run.sh" convert

echo "Running edit example..."
"$script_dir/run.sh" edit

echo "Running merge-split example..."
"$script_dir/run.sh" merge-split

echo "Running query example..."
"$script_dir/run.sh" query

echo "Running transfer example..."
"$script_dir/run.sh" transfer

echo "Running validation-cf example..."
"$script_dir/run.sh" validation-cf

echo "=== All tests and examples completed ==="
