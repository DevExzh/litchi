# Paired PPTX slide properties

This standalone harness compares one native primary slide-property edit with
one unified `File::set` call on generated one-slide PPTX corpora. Each sample
sets the authored slide name to `renamed` and visibility to hidden. The native
leg calls `Edit::set_slide_properties(Position::new(0), Some("renamed"), Some(true))`
and commits once. The unified
leg calls `File::set("/slide[1]", props)` once with `name` and `hidden`.

The default corpus sizes are `1,10,100,1000,10000` direct text-box shapes,
both eager and source owners, and nine ABBA rounds. Three rounds are the
minimum. Override values with:

```sh
unified-pptx-slide-properties-baseline --counts 1,10,1000 --owners eager,source --rounds 3
```

Run from the repository root with no concurrent build, test or measurement job:

```sh
BUILD=/tmp/litchi-unified-pptx-slide-properties-target
BIN="$BUILD/release/unified-pptx-slide-properties-baseline"
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR="$BUILD" cargo build --release \
  --manifest-path tools/perf-baseline/unified-pptx-slide-properties/Cargo.toml
LITCHI_SOURCE_COMMIT="$(git rev-parse HEAD)" \
LITCHI_RUSTC_VERSION="$(rustc --version)" \
LITCHI_TARGET="$(rustc -vV | sed -n 's/^host: //p')" \
LITCHI_BUILD_COMMAND='CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=/tmp/litchi-unified-pptx-slide-properties-target cargo build --release --manifest-path tools/perf-baseline/unified-pptx-slide-properties/Cargo.toml' \
taskset -c 0 "$BIN" --rounds 9 > /tmp/unified-pptx-slide-properties.json
```

The JSON schema is version `2`. Route names end in `slide_properties_v2` and
cover eager/source × native primary/unified file set. Every result records the
generated corpus byte count and SHA-256, the requested properties, exact or
decoded-member warm-up equivalence, raw ABBA nanosecond samples, upper-middle
medians, native combined-setter/one-commit accounting and unified one-call
accounting. Build provenance records the crate version and optional source,
compiler, target and build-command values. Failures are retained in a
structured `failure_matrix`, and the process exits nonzero after printing JSON
if any matrix case fails.

Corpus generation, opening, target parsing, property preparation, warm-up,
output serialization, equivalence checks and teardown are outside the timer.
This measures the requested native combined setter and unified set call; it
does not measure file-system save, allocation or RSS, real-world decks, or
complete goal acceptance. Never commit the compiled binary.
