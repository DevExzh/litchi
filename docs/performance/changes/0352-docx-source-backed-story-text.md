# Change 0352: DOCX source-backed selected-story text lifecycle

Status: correctness and CRUD closure only

`performance_claim: none`

## Selected-story scope

`SourceBackedPackage` now covers the bounded snapshot and text-streaming
lifecycle for one selected DOCX story: `Main`, `Header(index)`, or
`Footer(index)`. Eligible direct-paragraph text edits produce reversible
source-bound patches and inverses. Publication uses a same-topology one-part
overlay, with exact no-op output and trailing-byte copying preserved.

The resolver validates canonical relationships and content types, external
targets, shared targets, resolved namespaces, and markup-compatibility
processing. Ambiguous or unsupported XML is refused rather than approximated.
Freshness, source lineage, fingerprints, signature state, cancellation, and
failure atomicity remain enforced. Managed reads retain their typed boundary;
managed edits are refused with a typed error.

The focused coverage includes strict duplicate/end-tag validation, inverse
hostile-writer refusal, decoded namespaces, and actual emitted-byte bounds.

## Evidence and resource boundary

The final successful package/scenario-scoped commands, run serially, were:

```sh
ulimit -v 8388608
CARGO_TARGET_DIR=/home/zhuhe/CodeProjects/.cargo-targets/change-0352 \
CARGO_BUILD_JOBS=1 \
CARGO_INCREMENTAL=0 \
CARGO_PROFILE_DEV_DEBUG=0 \
CARGO_PROFILE_TEST_DEBUG=0 \
RUST_TEST_THREADS=1 \
cargo test -p litchi-docx --test source_backed_story_text -- --test-threads=1
# => 11/11

ulimit -v 8388608
CARGO_TARGET_DIR=/home/zhuhe/CodeProjects/.cargo-targets/change-0352 \
CARGO_BUILD_JOBS=1 \
CARGO_INCREMENTAL=0 \
CARGO_PROFILE_DEV_DEBUG=0 \
CARGO_PROFILE_TEST_DEBUG=0 \
RUST_TEST_THREADS=1 \
cargo test -p litchi-docx --test source_backed -- --test-threads=1
# => 16/16
```

The dedicated target reached 347 MiB. Post-run available memory was observed
at approximately 15 GiB while swap remained saturated; no Cargo processes ran
concurrently. These observations document the constrained run only and are
not a total-memory or host-resource bound.

## Remaining scope and claim boundary

Footnotes, endnotes, comments, and glossary stories are outside this slice.
Managed edits remain refused. No latency, throughput, RSS, allocation,
physical-I/O, streaming-speed, or benchmark claim follows, and no selector,
artifact, or whole-GOAL completion claim is added.
