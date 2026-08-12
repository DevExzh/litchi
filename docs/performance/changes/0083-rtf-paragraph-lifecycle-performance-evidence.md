# Change 0083: RTF paragraph lifecycle performance evidence

Date: 2026-08-12

Production capability: `c22a7303f5a4f01738cce64617f16d073a8bc1c3`

Status: selectable deterministic baseline; no comparative performance claim

## Scope and matched operations

This change adds harness, documentation, and CI coverage for the existing
native `litchi_rtf::edit::Edit::remove_paragraph` and `move_paragraph` APIs. It
does not change production RTF code or broaden their publication closure.

Both cases consume the identical immutable generated default-formatted
plain-RTF lifecycle corpus. This is distinct from the broader semantic RTF
read/edit corpus because that corpus deliberately carries explicit body font
formatting outside the lifecycle publisher's proven closure. For a
corpus of `N` paragraphs, the selectors are exact and deterministic:

- `rtf_semantic_remove_paragraph_save` removes that exact middle paragraph;
- `rtf_semantic_move_paragraph_save` moves source position zero to final
  position `N - 1`, where
  positions mean ordinals in the completed list as specified by the API.

Parsing, expected-output construction, sink allocation/reservation, durable
patch construction and the complete correctness oracles occur outside timing.
The
timed interval contains construction of one edit, one exact lifecycle staging
call, commit with the production candidate validation/reopen, a constant-size
diagnostics assertion, one shared snapshot-handle clone, and one
`Document::write_to` into a pre-reserved forward-only non-seek sink. The
diagnostics check is intentionally retained inside the common save-case branch;
the full byte, semantic, patch, durable and stale-source checks remain untimed.

These cases are matched workload coverage, not a control/candidate comparison.
No latency, instruction, allocation, peak-heap, RSS, or materialization
improvement is claimed. Any future comparative claim requires frozen identical
harnesses and a CPU-pinned balanced ABBA run with retained raw reports.

## Deterministic corpus and correctness oracle

The `litchi-rtf-paragraph-lifecycle-v1` generator supplies all shapes. Tiny has
24 ordinary body paragraphs, 1,304 source bytes, and SHA-256
`73641cf09b630632deabce8585c67f395a6bd3ac01eedcca6a8b7224ef00d252`.
Its removal output is 1,250 bytes with SHA-256
`49ef949a6ee85cc3a1bce19026e10a3b953136c73997eec6f719940e2c0b37a2`;
its reorder output remains 1,304 bytes with SHA-256
`9c7e42060e71be8cedf54fed9907d6a189efa45f1fec0f57d483e02af756f1fd`.

Large has 10,000 paragraphs, 540,008 source bytes, and SHA-256
`5feae6e5bd27751d75c936a5ae9f57e00d391d7759557e07c535227c72f3cf7c`.
Its removal output is 539,954 bytes with SHA-256
`946eda640eaa0b1df36e38264aa3ef8dd08d95986de0517baa186fae4da8f95d`;
its reorder output remains 540,008 bytes with SHA-256
`6f6a12a8f44580bd2a3b243306e49e185931079d3abe369b780fe82d4d111f27`.
Each report identifies the exact input and changed output SHA-256.

Before retained samples begin, the harness independently builds the expected
changed snapshot and runs the complete lifecycle oracle:

- reopen the serialized result through public `Document::from_bytes`;
- compare exact paragraph count and every paragraph string in order;
- compare the complete flattened body text;
- apply the source-bound in-memory patch and its inverse, checking exact bytes;
- encode the reversible patch to deterministic JSON, decode it, apply forward,
  and apply its inverse, again checking exact bytes;
- reject the durable patch on a separately committed stale source; and
- for move, commit an equal-position move and require unchanged diagnostics,
  shared snapshot identity, exact source bytes, and an empty durable patch.

Every timed iteration then compares the complete sink bytes with the expected
output, reopens and checks the same full semantic projection, and repeats exact
volatile forward/inverse checks outside the measured interval.

The sink reserves exactly the expected output length before timing and refuses
either total bytes or an individual write beyond that ceiling. The report
records accepted bytes, write calls, largest write, and output SHA-256. Native
RTF exposes no logical-Part materialization counter, so this record does not
invent one or borrow the OPC meaning of materialization.

## Capability boundary

Changed paragraph lifecycle publication remains plain-source-only. Focused
harness coverage stages both changed operations on raw CP-1252 and LZFu inputs
and requires `UnsupportedSource` plus exact original bytes. It separately
requires the producer-watermark removal and an opaque/formatted removal and
reorder to refuse without changing source bytes.

Equal-position move is honestly broader because it does not publish changed
bytes: plain, raw CP-1252, LZFu, producer-watermark, and opaque/formatted
snapshots all retain exact bytes and shared snapshot identity. The selectable
changed cases nevertheless reject every non-plain variant so result matrices
cannot imply changed transport support.

## Reproduction

Run the two matched large baselines together:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 \
  --semantic-shape large --rtf-variant plain \
  --case rtf_semantic_remove_paragraph_save,rtf_semantic_move_paragraph_save \
  --json target/perf/rtf-paragraph-lifecycle.json
```

For a future implementation comparison, freeze binaries with an identical
operation-specific harness and use a CPU-pinned balanced control A, candidate
A, candidate B, control B order. A direct removal-versus-reorder latency ratio
must not be described as an optimization result.

## Verification

The intended gates are:

```sh
cargo fmt --manifest-path tools/perf-baseline/Cargo.toml -- --check
cargo check --locked --manifest-path tools/perf-baseline/Cargo.toml
cargo test --locked --manifest-path tools/perf-baseline/Cargo.toml \
  semantic_rtf -- --nocapture
cargo clippy --locked --manifest-path tools/perf-baseline/Cargo.toml \
  --all-targets -- -D warnings
```

CI runs all four RTF transport/producer variants on tiny and asserts that only
the two plain lifecycle records exist, each has one sample, an output hash, and
positive bounded sink counters. Scheduled release CI repeats the same contract
for tiny and large with 15 samples per case.

Formatting/media-preserving lifecycle rewrites, non-ASCII and compressed
changed publication, malformed/security corpora, additional real producers,
filesystem-atomic save, allocation/memory counters, and comparative ABBA data
remain outside this change.
