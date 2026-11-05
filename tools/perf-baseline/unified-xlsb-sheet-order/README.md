# Paired XLSB sheet-order permutation

This standalone harness compares the native XLSB complete-catalog reorder with
one unified `File::move_to("/sheet[1]", Placement::End)` call. It generates
`N` worksheets named `Sheet00000` through `Sheet{N-1}`, with one distinct
numeric marker in each worksheet. The permutation is `[1, ..., N-1, 0]`, so
the first tab moves to the end while its worksheet payload and active identity
must follow it.

The default matrix uses `N = 2, 10, 100, 1000`, eager and source owners, and
nine ABBA rounds. Three rounds are the minimum. Corpus generation, opening,
permutation/path preparation, warm-up, output serialization and equivalence
checks are outside the clocks. Source promotion is inside every source-owner
sample on both legs.

Each sample owns one prepared native permutation vector and one parsed unified
path before timing. The timed native wrapper consumes the vector and the native
edit borrows its slice; the timed unified wrapper consumes the owned parsed path
and `Placement::End`. Neither leg clones or reparses its operation input inside
the timed call. The JSON records these ownership details explicitly alongside
the raw ABBA samples and medians.

Run from the repository root with no concurrent build, test or measurement job:

```sh
BUILD=/tmp/litchi-unified-xlsb-sheet-order-target
BIN="$BUILD/release/unified-xlsb-sheet-order-baseline"
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR="$BUILD" cargo build --release \
  --manifest-path tools/perf-baseline/unified-xlsb-sheet-order/Cargo.toml
LITCHI_SOURCE_COMMIT="$(git rev-parse HEAD)" \
LITCHI_RUSTC_VERSION="$(rustc --version)" \
LITCHI_TARGET="$(rustc -vV | sed -n 's/^host: //p')" \
LITCHI_SOURCE_DIGEST="$(git rev-parse HEAD^{tree})" \
LITCHI_MACHINE="taskset-cpu-0" \
LITCHI_BINARY_SHA256="$(sha256sum "$BIN" | awk '{print $1}')" \
LITCHI_BUILD_COMMAND='CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=/tmp/litchi-unified-xlsb-sheet-order-target cargo build --release --manifest-path tools/perf-baseline/unified-xlsb-sheet-order/Cargo.toml' \
taskset -c 0 "$BIN" --rounds 9 > /tmp/unified-xlsb-sheet-order.json
```

The matrix can be narrowed for a focused run:

```sh
unified-xlsb-sheet-order-baseline --counts 2,1000 --owners eager,source --rounds 3
```

The JSON contains source/build/binary provenance, corpus hashes, the complete
permutation, exact ZIP hashes, a per-member mismatch report when exact bytes
differ, active-tab identity/index checks, preserved-cell checks, raw samples,
medians and the failure matrix. A decoded member mismatch exits nonzero. Keep
compiled binaries and large captures outside the repository; retain bounded JSON
receipts under `docs/unified-ops/results/` with build and source provenance.
For a dirty build, record the changed source files and their SHA-256 values in
addition to the parent commit; a HEAD tree hash alone does not identify the build.
This is scoped evidence for one sheet-order operation; it makes no batch-scaling
claim.
