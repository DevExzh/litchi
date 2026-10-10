# Paired PPTX fanout publication

This standalone harness compares one native primary edit with one unified
`File::set("/slide[1]/*", props)` call on a generated one-slide PPTX. The
native leg prepares all replacement payloads once, builds borrowed 256-target
chunk views inside each timed sample, submits them through one
`set_shape_text_chunks` or `set_shape_property_chunks` call, and commits once.
The unified leg uses the wildcard selector once for the same text or mixed
text and `x` workload.

The default counts are `1,10,100,1000,10000`, both eager and source owners,
and nine ABBA rounds. Three rounds are the minimum. Override values with:

```sh
unified-pptx-fanout-publication-baseline \
  --counts 1,10,1000 --owners eager,source --rounds 3
```

Run from the repository root with no concurrent build, test or measurement
job. The harness prints one pretty JSON document to standard output:

```sh
BUILD=/tmp/litchi-unified-pptx-fanout-publication-target
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR="$BUILD" cargo build --release \
  --manifest-path tools/perf-baseline/unified-pptx-fanout-publication/Cargo.toml
LITCHI_SOURCE_COMMIT="$(git rev-parse HEAD)" \
LITCHI_RUSTC_VERSION="$(rustc --version)" \
LITCHI_TARGET="$(rustc -vV | sed -n 's/^host: //p')" \
LITCHI_BUILD_COMMAND='CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=/tmp/litchi-unified-pptx-fanout-publication-target cargo build --release --manifest-path tools/perf-baseline/unified-pptx-fanout-publication/Cargo.toml' \
taskset -c 0 "$BUILD/release/unified-pptx-fanout-publication-baseline" --rounds 3 \
  > /tmp/unified-pptx-fanout-publication.json
```

The JSON schema is version `1`. Route names end in
`fanout_publication_v2`; the previous `shape_query_fanout_v1` harness remains
available unchanged. Each result records the corpus byte count and SHA-256,
owner, operation, exact or decoded-member equivalence method, raw ABBA
nanosecond samples, upper-middle medians, native chunk/call/commit accounting,
and the unified set-call count. Build provenance records the crate version and
the optional environment values shown above. A failed case is recorded while
the rest of the matrix continues; the process exits nonzero after printing the
JSON if any case failed.

Corpus generation, opening, replacement/property preparation, warm-up,
output serialization, equivalence checks and teardown are outside the timer.
The small borrowed native chunk-view vector is constructed inside the native
timer. This measures publication work at one native call and one commit versus
one wildcard `File::set`; it does not measure file-system save, allocation or
RSS, real-world decks, or complete goal acceptance. Never commit the compiled
binary.
