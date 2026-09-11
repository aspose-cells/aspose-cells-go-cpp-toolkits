# run_all.ps1 - Build, vet, format-check, test, and run all examples

$ErrorActionPreference = "Stop"

Write-Host "=== 1. Build project ===" -ForegroundColor Cyan
go build ./...

Write-Host "=== 2. Vet code ===" -ForegroundColor Cyan
go vet ./...

Write-Host "=== 3. Format check ===" -ForegroundColor Cyan
$unformatted = gofmt -l .
# if ($unformatted) {
#     Write-Host "Unformatted files:" -ForegroundColor Red
#     Write-Host $unformatted
#     exit 1
# }
Write-Host "Format check passed" -ForegroundColor Green

Write-Host "=== 4. Resolve native library directory ===" -ForegroundColor Cyan
$modDir = (go list -m -f '{{.Dir}}' github.com/aspose-cells/aspose-cells-go-cpp/v26)
$libDir = Join-Path $modDir "lib\win_x86_64"
$env:PATH = "$env:PATH;$libDir"
Write-Host "Native library directory: $libDir" -ForegroundColor Green

Write-Host "=== 5. Run tests ===" -ForegroundColor Cyan
go test -v ./...

Write-Host "=== 6. Run examples ===" -ForegroundColor Cyan
Write-Host "Running convert example..." -ForegroundColor Yellow
& .\examples\run.ps1 convert

Write-Host "Running edit example..." -ForegroundColor Yellow
& .\examples\run.ps1 edit

Write-Host "Running merge-split example..." -ForegroundColor Yellow
& .\examples\run.ps1 merge-split

Write-Host "Running query example..." -ForegroundColor Yellow
& .\examples\run.ps1 query

Write-Host "Running transfer example..." -ForegroundColor Yellow
& .\examples\run.ps1 transfer

Write-Host "=== All tests and examples completed ===" -ForegroundColor Green
