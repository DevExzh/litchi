# Change 0121: native PPT repeated selected-query evidence

Date: 2026-08-15
Status: opt-in correctness and logical-read evidence; no latency claim

## Scope

The performance harness adds two opt-in, matched controls over the existing
deterministic native-PPT writer corpus:

- `ppt_semantic_repeated_shape_text` keeps one prepared eager presentation and
  performs eight identical selected-shape text queries;
- `ppt_source_backed_repeated_shape_text` keeps one prepared immutable
  source-backed snapshot and performs the same eight queries.

Corpus construction, source validation, and semantic-owner setup are outside
the elapsed interval. Warmups are isolated from measured samples. Every query
must return the exact deterministic text, and the report stores the SHA-256 of
the repeated semantic replay. Existing default case selection is unchanged;
these controls are opt-in.

## Evidence boundary

Source-backed timing uses an uninstrumented `OwnedSource`. A separate untimed
instrumented replay runs once per measured sample after source setup and
records `ReadAt` calls, bytes, and bytes in each later logical read that were
already covered by an earlier range (`source_read_range_overlap_bytes`). Range
evidence uses a sorted, coalesced interval union, so each current-read byte is
counted at most once (the same byte may contribute again on a later query);
tracking is enabled only for this replay and has a bounded interval budget.
The report also includes the selected target, canonical text hash, query count,
and one canonical repeated semantic hash (all measured samples are gated
against it). Eager controls retain the same semantic hash fields while
source-read vectors are empty.

These are logical reads through the generated fixture. They do not represent
physical disk I/O, decompression volume, allocation, RSS, or cache behavior.
The production regression
`repeated_source_reads_reuse_validated_cfb_index_and_reject_mutation` binds the
following exact two-query logical evidence to its deterministic fixture:

| implementation | `ReadAt` calls | bytes |
|---|---:|---:|
| legacy per-query CFB reconstruction | 74 | 8,310 |
| retained parsed CFB index | 66 | 3,190 |

The figures are call/byte counts only; they are not a latency or resource
claim.

## Claims deliberately not made

This change makes no latency, throughput, tail-latency, allocation, memory,
or release-performance claim. Any such claim requires a frozen release build,
matched eager/source controls, disclosed environment, retained raw samples,
and a controlled balanced ABBA run. The regression's exact call/byte delta is
reported above as logical-I/O evidence only.

The harness checks used for this change are:

```text
RUSTFLAGS='-D warnings -D deprecated' cargo check --locked -j1 \
  -p litchi-ole-common -p litchi-ppt --all-targets
cargo test --locked -j1 -p litchi-ppt --lib \
  repeated_source_reads_reuse_validated_cfb_index_and_reject_mutation
rustfmt --edition 2024 --check tools/perf-baseline/src/main.rs
CARGO_INCREMENTAL=0 cargo check --manifest-path tools/perf-baseline/Cargo.toml --offline -j1
CARGO_INCREMENTAL=0 cargo run --manifest-path tools/perf-baseline/Cargo.toml --offline -- \
  --case ppt_semantic_repeated_shape_text,ppt_source_backed_repeated_shape_text \
  --writer-shape tiny --semantic-shape tiny --samples 2 --warmup 1 --json report.json
```
