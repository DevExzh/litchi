# Paired XLSB blank-sheet insertion

This standalone harness compares one native `Edit::insert_sheet` operation
with one unified `File::add` operation. Each generated corpus contains `N`
worksheets named `Sheet00000` through `Sheet{N-1}`. Every source worksheet has
one distinct numeric marker. The operation inserts a blank, visible worksheet
named `Inserted` at the complete-catalog position `floor(N/2)`, so the marker
and active-tab checks detect changes to existing worksheet identity.

The default matrix uses `N = 1, 10, 100, 1000`, eager and source owners, and
nine ABBA rounds. Corpus generation, opening, position/node preparation,
warm-up, output serialization and equivalence checks are outside the clocks.
Source promotion is inside every source-owner sample on both legs.

Each sample owns a prepared insertion input before timing. The native input is
the checked position scalar and an owned name. The unified input is an owned parent path and
`NewNode` carrying the sheet name, visible state and positional placement. The
timed calls perform one native insertion or one unified add and no batch
operation is measured or claimed.

Run from the repository root with no concurrent build, test or measurement job:

```sh
BUILD=/tmp/litchi-unified-xlsb-sheet-insert-target
BIN="$BUILD/release/unified-xlsb-sheet-insert-baseline"
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR="$BUILD" cargo build --release \
  --manifest-path tools/perf-baseline/unified-xlsb-sheet-insert/Cargo.toml
LITCHI_SOURCE_COMMIT="$(git rev-parse HEAD)" \
LITCHI_RUSTC_VERSION="$(rustc --version)" \
LITCHI_TARGET="$(rustc -vV | sed -n 's/^host: //p')" \
LITCHI_SOURCE_DIGEST="$(git rev-parse HEAD^{tree})" \
LITCHI_MACHINE="taskset-cpu-0" \
LITCHI_BINARY_PATH="$BIN" \
LITCHI_BINARY_BYTES="$(stat -c '%s' "$BIN")" \
LITCHI_BINARY_SHA256="$(sha256sum "$BIN" | awk '{print $1}')" \
LITCHI_BUILD_COMMAND='CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=/tmp/litchi-unified-xlsb-sheet-insert-target cargo build --release --manifest-path tools/perf-baseline/unified-xlsb-sheet-insert/Cargo.toml' \
taskset -c 0 "$BIN" --rounds 9 > /tmp/unified-xlsb-sheet-insert.json
```

`LITCHI_SOURCE_MANIFEST` may contain a compact JSON array of the source files
used for a dirty build, with each entry recording its path, byte length and
SHA-256. The binary path, byte length and SHA-256 are copied into the JSON
receipt as a separate manifest.

The matrix can be narrowed for a focused run:

```sh
unified-xlsb-sheet-insert-baseline --counts 1,1000 --owners eager,source --rounds 3
```

The JSON contains source/build/binary provenance, corpus hashes, insertion
position and canonical path, exact ZIP hashes, a per-member mismatch report
when exact bytes differ, the blank inserted worksheet, preserved markers,
visible state, active-tab identity/index checks, raw samples, medians and a
failure matrix. A decoded member mismatch exits nonzero. Keep compiled
binaries and large captures outside the repository; retain bounded JSON
receipts under `docs/unified-ops/results/` with build and source provenance.
For a dirty build, record the changed source files and their SHA-256 values in
addition to the parent commit; a HEAD tree hash alone does not identify the
build.
