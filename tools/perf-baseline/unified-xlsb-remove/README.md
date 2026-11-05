# Paired XLSB cell-removal benchmark

This standalone harness compares package-native XLSB cell removal with the
unified operations routes. The native leg uses one
`Package::edit_cell_values(worksheet 0)` edit, removes `N` selected cells, and
commits once. The unified range leg uses one `File::remove` call for
`/sheet[1]/A1:A{N}`. For `N=1`, the matrix also records an explicit
`File::edit` removal transaction. The command-batch leg builds `N` `BatchItem` values before
the timer and executes them with `File::batch(..., BatchMode::Atomic)`, so the
batch admission and any native cell-removal coalescing are measured. For
`N=1`, the matrix also records the one-shot direct-cell route
`File::remove("/sheet[1]/A1")`.

The generated worksheet contains the selected numeric cells in column A and
one B1 sentinel that every output must retain. The default matrix measures
eager and source owners, nine ABBA rounds, and `1,10,100,1000,10000` selected
cells. Three rounds are the minimum. Override values with:

```sh
unified-xlsb-remove-baseline --counts 1,1000 --owners eager,source --rounds 9
```

Run from the repository root with no concurrent build, test or measurement job:

```sh
BUILD=/tmp/litchi-unified-xlsb-remove-target
BIN="$BUILD/release/unified-xlsb-remove-baseline"
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR="$BUILD" cargo build --release --locked \
  --manifest-path tools/perf-baseline/unified-xlsb-remove/Cargo.toml
LITCHI_SOURCE_COMMIT="$(git rev-parse HEAD)" \
LITCHI_SOURCE_DIGEST="$(git diff HEAD --binary | sha256sum | awk '{print $1}')" \
LITCHI_RUSTC_VERSION="$(rustc --version)" \
LITCHI_TARGET="$(rustc -vV | sed -n 's/^host: //p')" \
LITCHI_MACHINE="$(hostname)" \
LITCHI_BINARY_SHA256="$(sha256sum "$BIN" | awk '{print $1}')" \
LITCHI_BUILD_COMMAND='CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=/tmp/litchi-unified-xlsb-remove-target cargo build --release --locked --manifest-path tools/perf-baseline/unified-xlsb-remove/Cargo.toml' \
taskset -c 0 "$BIN" --rounds 9 > /tmp/unified-xlsb-remove.json
```

The JSON records source and binary provenance, corpus hashes, raw ABBA samples,
upper-middle medians, exact ZIP equivalence with decoded-member fallback,
native one-commit accounting, unified batch diagnostics, and a structured
`failure_matrix`. The successful default run has an empty failure matrix; a
failed case remains in the JSON and causes a nonzero exit status.

Opening, corpus generation, target and batch-item preparation, warm-up,
serialization, equivalence checks and teardown are outside the timer. Source
native samples include `SourceBackedWorkbook::to_workbook().into_package()` in
the native edit clock so the source owner boundary is comparable. The command
route reports whether the batch retained one native commit; the harness does
not claim command scaling unless that diagnostic is one commit. Never commit
the compiled binary or large raw captures; reduced JSON evidence must remain
below repository evidence limits.
