# Continuous integration

The toolkit is verified by two equivalent pipelines so that both hosted
platforms gate changes the same way:

| Platform | File | Runner |
| --- | --- | --- |
| GitHub Actions | [`.github/workflows/ci.yml`](../.github/workflows/ci.yml) | `ubuntu-latest`, `windows-latest` |
| GitLab CI | [`.gitlab-ci.yml`](../.gitlab-ci.yml) | `golang:1.21-bookworm` (Linux) |

## What runs

Both pipelines separate work that needs the native engine from work that does
not, because the engine is only usable when a license is configured (evaluation
mode has a 100-load quota per process and can corrupt worksheet state at load
time).

### Always runs

- `gofmt -l` — fails on any unformatted file.
- `go vet ./...`.
- `go build ./...`.
- `go test ./... -run '^$'` — compiles every test binary (including `TestMain`)
  without running test bodies. This still links the native library, so linkage
  problems surface on every change.

### Runs when a license is configured

- `go test ./... -count=1 -timeout 20m` — the full engine suite.
- `go test ./tests -race -count=1 -timeout 30m` — the race detector.
- Coverage over every package (`-coverpkg=./...`), with a **60% statement
  floor**. The pipeline fails below the floor and uploads `coverage.out` as an
  artifact.

Without a license the engine steps print a notice and are skipped; the static
and compile gates still run. This is what forks and public merge requests see.

## Configuring the license secret

The license is passed as a **base64-encoded file** in a variable named
`ASPOSE_LICENSE_B64`. Base64 keeps the value single-line and safe for both
platforms, and avoids depending on a file-type variable that only one platform
supports.

Create the value locally:

```sh
# Linux / macOS / Git Bash
base64 < /path/to/Aspose.Cells.lic | tr -d '\n'

# PowerShell
[Convert]::ToBase64String([IO.File]::ReadAllBytes("C:\path\to\Aspose.Cells.lic"))
```

### GitHub Actions

1. Open **Settings → Secrets and variables → Actions → New repository secret**.
2. Name it `ASPOSE_LICENSE_B64` and paste the base64 value.

Repository secrets are not exposed to pull requests from forks, which is why
the license steps are conditional rather than required.

### GitLab CI

1. Open **Settings → CI/CD → Variables → Add variable**.
2. Key: `ASPOSE_LICENSE_B64`; type: **Variable** (not File); mask and protect it.

The GitLab test jobs only exist when the variable is present
(`rules: if: $ASPOSE_LICENSE_B64`). For merge requests from forks, protect the
variable so it is not exposed; the quality job still runs.

## Native library resolution

The Aspose.Cells Go module ships its native library inside the module cache:

- Linux: `<GOMODCACHE>/github.com/aspose-cells/aspose-cells-go-cpp/v26@v26.10.0/lib/linux_x86_64`
- Windows: `<GOMODCACHE>/github.com/aspose-cells/aspose-cells-go-cpp/v26@v26.10.0/lib/win_x86_64`

The pipelines add that directory to `LD_LIBRARY_PATH` (Linux) or `PATH`
(Windows) after `go mod download`, so no manual library installation is needed.

**Windows** additionally needs a MinGW toolchain for cgo. The workflow checks
for `gcc` and installs the Chocolatey `mingw` package when it is missing.
**Linux** runners already provide `gcc`/`g++`.

## Running the same gates locally

Linux:

```sh
export LD_LIBRARY_PATH="$(go env GOMODCACHE)/github.com/aspose-cells/aspose-cells-go-cpp/v26@v26.10.0/lib/linux_x86_64:$LD_LIBRARY_PATH"

git ls-files '*.go' | grep -v '^agent/ResultOutput/' | xargs gofmt -l
go vet ./...
go build ./...
go test ./... -run '^$' -count=1

# With a license:
export LicenseFilePath=/path/to/Aspose.Cells.lic
go test ./... -count=1 -timeout 20m
go test ./tests -race -count=1 -timeout 30m
go test ./tests -count=1 -coverpkg=./... -coverprofile=coverage.out -timeout 20m
go tool cover -func=coverage.out | tail -1
```

Windows (Git Bash). Exporting a POSIX `PATH` is not enough for the Windows
DLL loader, so set `PATH` inside `cmd.exe` for the test process:

```sh
lib="$(go env GOMODCACHE)/github.com/aspose-cells/aspose-cells-go-cpp/v26@v26.10.0/lib/win_x86_64"
cmd //c "set PATH=$(cygpath -w "$lib");%PATH% && go test ./... -count=1 -timeout 20m"
```

Alternatively, `go test -c -o out/tests.exe ./tests`, copy both DLLs beside
`tests.exe`, and run it directly. PowerShell:

```powershell
$lib = "$(go env GOMODCACHE)\github.com\aspose-cells\aspose-cells-go-cpp\v26@v26.10.0\lib\win_x86_64"
$env:Path = "$lib;$env:Path"
go test ./... -count=1 -timeout 20m
```

The coverage floor is enforced by CI, not by `go test`; compare the `total:`
line yourself when running locally.

## Why the license is optional in CI

Evaluation mode is a supported way to use the toolkit, and requiring a secret
would make the pipelines unusable for forks. Skipping only the engine steps
keeps the most valuable checks — formatting, vet, compilation, and native
linkage — available to everyone while keeping licensed runs deterministic and
quota-free.
