# Paired PPTX shape removal

This standalone harness compares one native primary shape removal with one
unified `File::remove` call on a generated one-slide PPTX. Each sample removes
the last direct shape: native `Edit::remove_shape(Position::new(0),
count - 1)` followed by one commit, or unified
`File::remove("/slide[1]/shape[N]")` using the corresponding one-based path.

The default corpus sizes are `1,10,100,1000,10000`, both eager and source
owners, and nine ABBA rounds. Three rounds are the minimum. Override values
with:

```sh
unified-pptx-shape-removal-baseline --counts 1,10,1000 --owners eager,source --rounds 3
```

Run from the repository root with no concurrent build, test or measurement job:

```sh
BUILD=/tmp/litchi-unified-pptx-shape-removal-target
BIN="$BUILD/release/unified-pptx-shape-removal-baseline"
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR="$BUILD" cargo build --release \
  --manifest-path tools/perf-baseline/unified-pptx-shape-removal/Cargo.toml
LITCHI_SOURCE_COMMIT="$(git rev-parse HEAD)" \
LITCHI_RUSTC_VERSION="$(rustc --version)" \
LITCHI_TARGET="$(rustc -vV | sed -n 's/^host: //p')" \
LITCHI_BUILD_COMMAND='CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=/tmp/litchi-unified-pptx-shape-removal-target cargo build --release --manifest-path tools/perf-baseline/unified-pptx-shape-removal/Cargo.toml' \
taskset -c 0 "$BIN" --rounds 9 > /tmp/unified-pptx-shape-removal.json
```

The JSON schema is version `1`. Route names end in `shape_removal_v1` and
cover eager/source × native primary/unified file removal. Every result records
the generated corpus byte count and SHA-256, the zero-based native shape
position and one-based unified target, exact or decoded-member warm-up
equivalence, raw ABBA nanosecond samples, upper-middle medians, native
one-call/one-commit accounting and the unified one-call accounting. Build
provenance records the crate version and the optional environment values shown
above. Failures are retained in a structured `failure_matrix`, and the process
exits nonzero after printing JSON if any matrix case fails.

Each case reuses its immutable native and unified base owners after warm-up.
These are warmed-handle timings, not cold first-operation measurements.
Corpus generation, opening, target parsing, warm-up, output serialization,
equivalence checks and teardown are outside the timer. Timed outputs are held
until after each timer stops. This measures one native call and commit against
one unified removal call; it does not measure file-system save, allocation or
RSS, real-world decks, or complete goal acceptance. Never commit the compiled
binary.
