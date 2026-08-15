# Change 0139: ODP repeated-text cache evidence harness

## Decision

Change 0139 adds two opt-in performance-evidence selectors for the committed
ODP repeated-text cache. It changes only the harness and its documentation;
it does not change production ODP behavior. The selectors are:

- `odp_source_backed_repeated_text_uncached`
- `odp_source_backed_repeated_text_cached`

They bring the selectable `Case` matrix to 265 names while preserving the
default 36 cases / 198 records. This change accepts the matched correctness
and source-replay gates only. It makes no latency, physical-I/O,
decompression, allocation, RSS, cold-cache, or release claim until a frozen,
CPU-pinned measured ABBA run.

## Matched corpus and owner

Both selectors use the existing deterministic media-rich ODP generator
`litchi-odp-media-textbox-publication-v1` and the same source-backed owner
shape. The corpus is fixed at:

```text
slides: 12
archive members: 13
Pictures members: 8
uncompressed Pictures payload: 16,777,216 B (8 x 2 MiB)
source archive: 16,786,129 B
source archive SHA-256: c5e98dac88846d7b8264f0af4e893d80e21672222c35c3b8890f78cff53242d3
canonical full-text SHA-256: 460bfe509d9c35eb05728c4ff847e0a080aec9bf7a2684ee80b2f9e46b37e3c7
uncompressed Pictures payload SHA-256: bac87991b97be1a282eabbe32c245dc504bd4344aa01c6d0619b00d41f63983c
```

Corpus creation, archive topology, media ranges, source owner construction,
expected projections, output-slot reservation, and all validation are outside
the named timer. The summary keeps archive-member count, picture count,
aggregate uncompressed media bytes, archive/text/media hashes, and source
counter scopes together so an evidence record cannot be mistaken for a
generic ODP timing result.

## Matched timed work

Each measured sample prepares one `SourceBackedPresentation` owner and four
output slots before starting the timer. The timer contains exactly four full
text projections:

- The control reconstructs the pre-cache public sequence with
  `slides()`, `Slide::all_text()` for every slide, filtering empty text and
  joining the remaining values with the exact `\n\n` separator, followed by
  the trailing `check_source()` performed by the historical uncached path.
- The candidate calls `SourceBackedPresentation::text()` four times,
  exercising the production threshold-two cache.

After timing, both paths compare every projection to the eager oracle, check
the repeated-output digest, re-check the complete slide projection and source
freshness, and verify archive topology and media preservation. None of those
checks contributes to elapsed time. The control intentionally retains the
old public helper's allocation shape as the pre-cache oracle; it does not add
new control-only reservation work to the timed interval.

## Source replay and freshness gates

Each measured sample receives a separate untimed `InstrumentedSource` replay.
Preparation counters are recorded, then source counters are reset before the
four projections. The post-preparation replay must have all of the following
equal to zero for both selectors:

```text
source_replay_read_calls
source_replay_read_bytes
source_replay_range_overlap_bytes
source_replay_payload_read_calls
source_replay_payload_read_bytes
```

The replay also records source-version observations for every projection and
binds the matched freshness work:

```text
uncached control per-call vector: [3, 3, 3, 3]
cached candidate per-call vector: [3, 5, 2, 2]
aggregate observations per selector: 12
```

The candidate's extra observations on its second call are the cache
publication/freshness checks; later calls are cache hits with two observations
each. The replay captures both the aggregate vector and the per-call vector
in `source.odp_repeated_text`. Preparation payload counters are retained as
evidence rather than forced to zero: instrumented ZIP-tail preparation can
physically overlap a compressed `Pictures` range without materializing that
media payload. Only the post-preparation four-call replay is the zero-read
gate.

Archive topology and `verify_odp_media_archive` run outside timing for corpus
preparation, every measured sample, every instrumented replay, and final
validation. Thus the source replay cannot silently trade media preservation
for a text-only projection.

## Verification

The focused harness test is:

```text
cargo test --release --locked --manifest-path tools/perf-baseline/Cargo.toml \
  media_rich_odp_repeated_text_selectors_are_matched_and_source_fresh -- \
  --nocapture
```

The one-test run passes for both selectors and two samples. A two-sample
selector smoke also passes with the four-call replay gates and records the
corpus values above. The smoke is intentionally not release evidence: it is
not CPU-pinned, does not use a clean detached release binary, and does not
run balanced ABBA legs. The release decision remains deferred until those
conditions are met.

## Applicability and limitations

This evidence applies only to the two named selectors, the deterministic
media-rich ODP corpus, the source-backed owner shape, and the four repeated
full-text projections described here. It does not generalize to other ODP
documents, slide counts, source implementations, cold-cache behavior,
single-call text queries, allocations, memory, physical I/O, or unrelated
ODF/iWork paths. The implementation and schema live in
[`tools/perf-baseline/src/main.rs`](../../../tools/perf-baseline/src/main.rs),
with the user-facing selector contract in
[`tools/perf-baseline/README.md`](../../../tools/perf-baseline/README.md).
