#!/usr/bin/env bash
# Build litchi-py with cargo, stage the extension in a temporary directory and
# run the Python read-API oracle in this directory with the standard library's
# unittest. No maturin, pip or pytest is needed.
#
# Usage:
#   crates/pyo3-litchi/tests/run_tests.sh [--release] [--no-build] [unittest options]
#
#   --release    build and test the release profile (default: dev)
#   --no-build   test the library already in the target directory
#   anything else is passed to "python -m unittest discover", e.g. -v or -k d1
#
# Environment:
#   CARGO_TARGET_DIR  honoured; the library is looked up where cargo put it
#   PYTHON            interpreter to test with (default: python3)
#   LITCHI_TEST_DATA  fixture root (default: <repo>/test-data)
#
# Nothing is written into the source tree: the module is copied to a mktemp
# directory that is removed on exit, and bytecode caching is disabled. Build
# output goes wherever cargo sends it (CARGO_TARGET_DIR, else <repo>/target).
set -euo pipefail

tests_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
repo_root="$(cd "$tests_dir/../../.." && pwd)"
python="${PYTHON:-python3}"
profile="debug"
cargo_args=()
build=1

while [ $# -gt 0 ]; do
  case "$1" in
    --release) profile="release"; cargo_args=(--release); shift ;;
    --no-build) build=0; shift ;;
    --) shift; break ;;
    *) break ;;
  esac
done

if [ "$build" = 1 ]; then
  (cd "$repo_root" && cargo build -p litchi-py ${cargo_args[@]+"${cargo_args[@]}"})
fi

# cargo metadata resolves CARGO_TARGET_DIR and any build.target-dir setting.
target_dir="$(cd "$repo_root" && cargo metadata --format-version 1 --no-deps |
  "$python" -c 'import json, sys; print(json.load(sys.stdin)["target_directory"])')"

case "$(uname -s)" in
  Darwin) library="liblitchi_py.dylib" ;;
  *) library="liblitchi_py.so" ;;
esac
artifact="$target_dir/$profile/$library"
if [ ! -f "$artifact" ]; then
  echo "run_tests.sh: $artifact not found (build it, or drop --no-build)" >&2
  exit 1
fi

stage="$(mktemp -d "${TMPDIR:-/tmp}/litchi-py-oracle.XXXXXX")"
trap 'rm -rf "$stage"' EXIT
# The module name comes from #[pymodule] fn litchi_py; abi3 suffix as maturin
# would install it.
cp "$artifact" "$stage/litchi_py.abi3.so"

echo "run_tests.sh: $("$python" -c 'import sys; print(sys.version.split()[0])') with $artifact ($profile)"
status=0
(
  cd "$stage"
  PYTHONPATH="$stage${PYTHONPATH:+:$PYTHONPATH}" PYTHONDONTWRITEBYTECODE=1 \
    "$python" -m unittest discover -s "$tests_dir" -t "$tests_dir" -p 'test*.py' "$@"
) || status=$?
exit "$status"
