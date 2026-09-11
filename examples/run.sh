#!/usr/bin/env bash
# Run the example commands.
#
# Usage:
#   ./examples/run.sh            # run every example in order
#   ./examples/run.sh convert    # run a single example (convert|edit|merge-split|query|transfer)
#
# Each example is a self-contained command; it reads its sample data from
# examples/data and writes outputs to examples/<name>/out (absolute paths, so
# the same output lands there whether it is launched from the module root or
# from inside an example directory).
set -euo pipefail

cd "$(dirname "$0")/.."

run() {
  echo "==> examples/$1"
  (cd "examples/$1" && go run .)
}

case "${1:-all}" in
  all)
    run convert
    run edit
    run merge-split
    run query
    run transfer
    ;;
  convert | edit | merge-split | query | transfer)
    run "$1"
    ;;
  *)
    echo "unknown example: $1" >&2
    echo "usage: $0 [all|convert|edit|merge-split|query|transfer]" >&2
    exit 1
    ;;
esac
