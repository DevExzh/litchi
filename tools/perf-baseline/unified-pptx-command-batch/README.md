# Paired PPTX direct command batches

This standalone harness measures homogeneous PPTX shape updates through two
paired APIs.  The native leg creates one eager or source-backed primary `Edit`,
submits 256-target `set_shape_texts` or `set_shape_properties` chunks, and
commits once.  The unified leg submits one `File::batch(items,
BatchMode::Atomic)` call containing one distinct direct-path `set` command per
shape.  Text and mixed text-plus-`x` property workloads run for both owner
modes.

The default counts are `1,10,100,1000,10000`, with nine ABBA rounds.  Counts
and rounds are configurable. Both owners run by default; `--owners eager` or
`--owners source` selects one owner for a focused rerun:

```sh
unified-pptx-command-batch-baseline --counts 1,10,1000 --rounds 9
unified-pptx-command-batch-baseline --counts 10000 --owners eager --rounds 3
```

Build and run from the repository root with no concurrent Cargo, test or
measurement job:

```sh
BUILD=/tmp/litchi-unified-pptx-command-batch-target
BIN="$BUILD/release/unified-pptx-command-batch-baseline"
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR="$BUILD" cargo build --release \
  --manifest-path tools/perf-baseline/unified-pptx-command-batch/Cargo.toml
LITCHI_SOURCE_COMMIT="$(git rev-parse HEAD)" \
LITCHI_RUSTC_VERSION="$(rustc --version)" \
LITCHI_TARGET="$(rustc -vV | sed -n 's/^host: //p')" \
LITCHI_BUILD_COMMAND='CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=/tmp/litchi-unified-pptx-command-batch-target cargo build --release --manifest-path tools/perf-baseline/unified-pptx-command-batch/Cargo.toml' \
taskset -c 0 "$BIN" --rounds 9 > /tmp/unified-pptx-command-batch.json
```

The JSON schema is version `1`.  The eight routes are eager/source × native
primary/unified file batch × text/text+x, and every route name ends in the new
`command_batch_v1` identity.  Each result records the generated corpus
size and SHA-256, the direct path pattern, native chunk/call/commit counts,
unified operation and commit diagnostics, ABBA nanosecond samples, upper
middle medians and the unified/native ratio.  Build provenance is recorded
from the optional environment variables in the command above.

The harness requires every homogeneous direct-shape command batch to report
one unified native commit while retaining one applied result per command.  It
does not assume a particular native leaf-operation count; that diagnostic is
recorded for inspection.  A failure of the coalescing assertion stops the run
instead of presenting per-command commits as evidence for the batch scaling
route.

All generated command vectors, property values and native replacement vectors
are prepared before timing.  Opening, initial-owner checks, warm-up,
serialization, exact/decoded-member equivalence checks and destruction are
outside timed intervals.  Timed outputs remain retained until their timer
stops.  The harness excludes filesystem save, allocation/RSS, real-world
decks and full performance acceptance, and never commits its compiled binary.

Progress is printed to stderr outside timed regions. If a scenario fails, the
harness records its owner, operation, count, corpus identity and error in the
JSON `failures` array, continues the other scenarios, then exits nonzero. The
completed measurements remain available; an unsuccessful matrix is never
reported as complete acceptance.
