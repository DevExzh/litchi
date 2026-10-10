# Paired PPTX shape-query fanout

This standalone harness measures one selector fanout on a generated one-slide
PPTX.  It compares the native primary edit, which reuses one eager or
source-backed `Edit` and submits 256-shape `set_shape_texts` or
`set_shape_properties` chunks before one commit, with one unified
`File::set("/slide[1]/*", props)` call.  The wildcard is the public selector
used by the PPTX fanout tests and selects the slide's direct shapes.

The default counts are `1,10,100,1000,10000` and the default is nine ABBA
rounds.  Every round schedules native, unified, unified, native, producing two
samples per leg.  Override either value with:

```sh
unified-pptx-shape-fanout-baseline --counts 1,10,1000 --rounds 9
```

Run from the repository root, with no concurrent build, test or measurement
job.  The harness prints one pretty JSON document to standard output:

```sh
BUILD=/tmp/litchi-unified-pptx-shape-fanout-target
BIN="$BUILD/release/unified-pptx-shape-fanout-baseline"
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR="$BUILD" cargo build --release \
  --manifest-path tools/perf-baseline/unified-pptx-shape-fanout/Cargo.toml
LITCHI_SOURCE_COMMIT="$(git rev-parse HEAD)" \
LITCHI_RUSTC_VERSION="$(rustc --version)" \
LITCHI_TARGET="$(rustc -vV | sed -n 's/^host: //p')" \
LITCHI_BUILD_COMMAND='CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=/tmp/litchi-unified-pptx-shape-fanout-target cargo build --release --manifest-path tools/perf-baseline/unified-pptx-shape-fanout/Cargo.toml' \
taskset -c 0 "$BIN" --rounds 9 > /tmp/unified-pptx-shape-fanout.json
```

The JSON schema is version `1`.  Top-level `routes` contains the eight new
route names: eager/source × native-primary/unified-file-set × text/text+x.
`results` has one object per owner, operation and count.  Each result includes
the generated corpus byte count and SHA-256 under `artifact`, the selector,
chunk size, exact or decoded-member equivalence method, ABBA nanosecond sample
arrays, upper-middle medians and `unified_over_native`.  The unified direct
`File::set` API does not expose explicit transaction diagnostics; its result
therefore records `unified_set_calls: 1` and `unified_diagnostics: null`.
`build_provenance` records the crate version plus the optional environment
values shown in the command.  Route names and corpus identities are part of
the machine-readable record; changing their semantics requires new names.

The text operation sets visible text.  The mixed operation sets the same text
and moves every shape's `x` to `1in`; the native complete bounds are chosen to
match that public partial property update.  All target vectors and unified
properties are prepared before timing.  Opening, initial-owner checks,
warm-up, output serialization, exact/decoded output comparison, and teardown
are outside the timer.  Timed outputs are retained until after each timer
stops.  This harness does not measure file-system save, command batches,
allocation/RSS, real-world decks or full performance acceptance, and it never
commits a compiled binary.
