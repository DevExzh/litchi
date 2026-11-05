# Paired PPTX chunk publication

This harness compares one native primary edit using `set_shape_text_chunks`
or `set_shape_property_chunks` against one atomic `File::batch` with a separate
direct-path command per shape. Native 256-target chunks are admitted together
and published once. Eager and source owners run text and mixed text/position
workloads. New route names end in `chunk_publication_v2`; the previous
`command_batch_v1` harness and routes remain available unchanged.

Defaults are counts `1,10,100,1000,10000`, both owners, and nine ABBA rounds.
`--counts`, `--owners eager,source`, and `--rounds` permit focused runs.
Three rounds are the minimum, yielding six samples per leg. All command vectors
and replacement payloads are prepared before timing. The small borrowed native
chunk-view vector is constructed inside timing, along with native edit creation,
staging and commit. Unified timing covers its equivalent atomic batch call.
Generation, opening, initial checks, warm-up, output serialization, equivalence
checks and destruction are outside timing. Warm-up outputs must match exact
bytes or decoded ZIP members; every unified sample must report one native commit
and one changed result per command.

```sh
CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_INCREMENTAL=0 CARGO_TARGET_DIR=/tmp/litchi-unified-chunk-publication-target cargo build --offline --release --manifest-path tools/perf-baseline/unified-pptx-chunk-publication/Cargo.toml
taskset -c 0 /tmp/litchi-unified-chunk-publication-target/release/unified-pptx-chunk-publication-baseline --rounds 3 > /tmp/pptx-chunk-publication.json
```

Set `LITCHI_SOURCE_COMMIT`, `LITCHI_RUSTC_VERSION`, `LITCHI_TARGET`, and
`LITCHI_BUILD_COMMAND` to record build provenance. JSON schema 1 includes corpus
SHA-256, raw ABBA samples, upper-middle medians, ratios, native calls/chunks,
unified diagnostics and structured failures. A failed scenario is recorded and
the other scenarios continue; the process exits nonzero if any scenario failed.
Run without concurrent builds/tests. These timings do not establish allocation,
RSS, real-world corpus, save, or complete goal acceptance. Never commit binaries.
