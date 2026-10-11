# Paired XLSB worksheet removal

This standalone harness compares one native `Edit::remove_sheet` operation
with one unified `File::remove` operation. Each generated corpus contains
`N` worksheets named `Sheet00000` through `Sheet{N-1}`. Every worksheet owns a
distinct numeric marker. The operation removes the source worksheet at
`floor(N/2)`, leaving `N - 1` worksheets and preserving the active
`Sheet00000` identity.

The default matrix uses `N = 2, 10, 100, 1000`, eager and source owners, and
nine ABBA rounds. Corpus generation, opening, source position/path
preparation, warm-up, output serialization and equivalence checks are outside
the clocks. Source promotion is inside every source-owner sample on both
legs. Each timed sample performs one native removal and one commit, or one
unified `File::remove` call; this harness makes no batch-scaling claim.

Run from the repository root with no concurrent build, test or measurement
job:

```sh
BUILD=/tmp/litchi-unified-xlsb-sheet-remove-target
BIN="$BUILD/release/unified-xlsb-sheet-remove-baseline"
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR="$BUILD" cargo build --release --locked \
  --manifest-path tools/perf-baseline/unified-xlsb-sheet-remove/Cargo.toml
LITCHI_SOURCE_COMMIT="$(git rev-parse HEAD)" \
LITCHI_RUSTC_VERSION="$(rustc --version)" \
LITCHI_TARGET="$(rustc -vV | sed -n 's/^host: //p')" \
LITCHI_SOURCE_DIGEST="$(git rev-parse HEAD^{tree})" \
LITCHI_MACHINE="taskset-cpu-0" \
LITCHI_BINARY_PATH="$BIN" \
LITCHI_BINARY_BYTES="$(stat -c '%s' "$BIN")" \
LITCHI_BINARY_SHA256="$(sha256sum "$BIN" | awk '{print $1}')" \
LITCHI_BUILD_COMMAND='CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=/tmp/litchi-unified-xlsb-sheet-remove-target cargo build --release --locked --manifest-path tools/perf-baseline/unified-xlsb-sheet-remove/Cargo.toml' \
taskset -c 0 "$BIN" --rounds 9 > /tmp/unified-xlsb-sheet-remove.json
```

`LITCHI_SOURCE_MANIFEST` may contain a compact JSON array describing the
source files used for a dirty build. Each entry records its path, byte length
and SHA-256. The binary path, byte length and SHA-256 are recorded separately.

The matrix can be narrowed for a focused run:

```sh
unified-xlsb-sheet-remove-baseline --counts 2,1000 --owners eager,source --rounds 3
```

The JSON contains source/build/binary provenance, corpus hashes, the removed
source position and name, exact ZIP hashes, a decoded-member mismatch report
when exact bytes differ, surviving marker and active-tab checks, deleted
worksheet and binary-index member checks, raw samples, medians and a failure
matrix. A decoded member mismatch exits nonzero. Keep compiled binaries and
large captures outside the repository; retain only bounded JSON receipts under
`docs/unified-ops/results/`.
