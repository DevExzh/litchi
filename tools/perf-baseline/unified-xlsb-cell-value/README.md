# Paired XLSB cell-value benchmark

This standalone harness compares the package-owned XLSB cell-value edit with
the unified facade. It generates one worksheet containing row-major numeric
cells, then replaces those values with `Package::edit_cell_values` plus one
`Edit::set_scalar` call per cell, or with the unified `File` API. The native
leg always uses one package edit. The unified one-shot case uses
`File::set("/sheet[1]/A1", {"value": number})`; the 1/10/100/1,000/10,000
cell cases use one `File::edit` transaction with repeated direct cell sets so
the result exposes whether the adapter coalesced them into one leaf commit.

The default matrix measures eager and source owners, 3 ABBA rounds, six
samples per leg, and 1, 10, 100, 1,000 and 10,000 staged cells. For the
one-cell corpus it records both the one-shot `set` and the transaction `edit`
shape. Opening, corpus generation, target/property preparation, warm-up,
serialization and equivalence checks are outside the clocks. Source-owner
native samples include `SourceBackedWorkbook::to_workbook().into_package()`
inside the timed native edit, matching the unified source edit boundary.
The eager native leg starts from `Package`, while the eager facade starts from
`Workbook`; this is a package-native comparison with different prepared owner
types. Facade property recollection, admission and coercion are timed even
though caller target/property preparation is outside the clocks.

Run from the repository root with no concurrent build, test or measurement job:

```sh
BUILD=/tmp/litchi-unified-xlsb-cell-value-target
BIN="$BUILD/release/unified-xlsb-cell-value-baseline"
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR="$BUILD" cargo build --release \
  --manifest-path tools/perf-baseline/unified-xlsb-cell-value/Cargo.toml
LITCHI_SOURCE_COMMIT="$(git rev-parse HEAD)" \
LITCHI_SOURCE_DIGEST="$(git diff HEAD --binary | sha256sum | awk '{print $1}')" \
LITCHI_RUSTC_VERSION="$(rustc --version)" \
LITCHI_TARGET="$(rustc -vV | sed -n 's/^host: //p')" \
LITCHI_MACHINE="$(hostname)" \
LITCHI_BINARY_SHA256="$(sha256sum "$BIN" | awk '{print $1}')" \
LITCHI_BUILD_COMMAND='CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=/tmp/litchi-unified-xlsb-cell-value-target cargo build --release --manifest-path tools/perf-baseline/unified-xlsb-cell-value/Cargo.toml' \
taskset -c 0 "$BIN" > /tmp/unified-xlsb-cell-value.json
```

The matrix can be narrowed for a smoke measurement:

```sh
unified-xlsb-cell-value-baseline --counts 1,1000 --owners eager,source --rounds 3
```

The JSON records the corpus and source digests, build and machine provenance,
raw ABBA nanosecond samples, upper-middle medians, ratios, native operation
counts, unified transaction diagnostics, and warm-up output equivalence. Exact
ZIP bytes are preferred. If bytes differ, every decoded ZIP member is compared
and the report identifies each missing or differing member with its byte length
and SHA-256; a mismatch exits nonzero. The result is scoped measurement
evidence, not a claim about the complete unified-operations goal. Never commit
the compiled binary. Keep reduced measurement evidence below repository size
limits and retain larger raw captures outside the checkout. The source digest
command above covers tracked modifications; include new source files in a
separate manifest when measuring before the implementation commit.
