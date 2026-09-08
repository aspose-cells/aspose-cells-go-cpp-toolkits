# run.ps1 - Run the example commands.
#
# Usage:
#   powershell examples/run.ps1            # run every example in order
#   powershell examples/run.ps1 convert    # run a single example (convert|edit|merge-split|query|transfer)
#
# Each example is a self-contained command; it reads its sample data from
# examples/data and writes outputs to examples/<name>/out (absolute paths, so
# the same output lands there whether it is launched from the module root or
# from inside an example directory.

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location (Join-Path $scriptDir "..")

function Run-Example {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    Write-Host "==> examples/$Name" -ForegroundColor Cyan
    $examplePath = Join-Path $scriptDir $Name
    Set-Location $examplePath
    go run .
}

# Main
if ($args.Count -eq 0) {
    $target = "all"
} else {
    $target = $args[0]
}

switch ($target) {
    "all" {
        Run-Example "convert"
        Run-Example "edit"
        Run-Example "merge-split"
        Run-Example "query"
        Run-Example "transfer"
    }
    {$_ -eq "convert" -or $_ -eq "edit" -or $_ -eq "merge-split" -or $_ -eq "query" -or $_ -eq "transfer"}  {
        Run-Example $target
    }
    default {
        Write-Host "unknown example: $target" -ForegroundColor Red
        Write-Host "usage: powershell examples/run.ps1 [all|convert|edit|merge-split|query|transfer]" -ForegroundColor Red
        exit 1
    }
}
Set-Location (Join-Path $scriptDir "..")