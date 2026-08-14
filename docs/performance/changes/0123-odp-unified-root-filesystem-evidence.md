# Change 0123: ODP unified-root filesystem evidence

Date: 2026-08-15
Status: opt-in correctness and logical-range evidence; no latency claim

## Scope

The standalone harness adds four matched selectors over the existing
deterministic 12-slide/eight-2 MiB `Pictures/` ODP corpus:

| eager byte-backed root | filesystem source-backed root |
|---|---|
| `odp_file_eager_open` | `odp_file_source_open` |
| `odp_file_eager_selected_slide` | `odp_file_source_selected_slide` |

The eager controls use the unified `litchi::Presentation` root with a prepared
byte buffer. The source controls call `litchi::Presentation::open` on one
temporary corpus file, exercising the production ODP filesystem handoff. The
four selectors are opt-in: the matrix grows from 233 to 237 names while the
default matrix remains 36 cases / 198 records.

## Timing and parity

Corpus generation and temporary-file creation/writing happen before timing.
Open selectors time only matching root owner construction: eager
`Presentation::from_bytes` versus source `Presentation::open`. Selected-slide
selectors prepare their corresponding owner before timing and time only the
middle-slide query. A byte-buffer clone is prepared before the eager-open timer
so it does not become an accidental source-path fairness cost.

After every measured operation, the harness verifies the complete root slide
projection and canonical semantic digest, selected-slide text, metadata digest,
source-file byte count and SHA-256, and exact archive/member identity. It then
reopens the fixed typed ODP owner outside timing and verifies all slide/media
semantics plus the selected uncompressed media payload and digest. Thus eager
and filesystem controls carry matched semantic, metadata, media, member, and
hash gates without extending the measured interval.

## Independent source replay

Each measured source sample has a separate direct
`litchi_odp::SourceBackedPresentation` replay over an instrumented positional
source. The replay records source calls/bytes, prior-range overlap, aggregate
compressed `Pictures/*` overlap, and selected-media calls/bytes. Open replay
captures catalog reads; the middle-slide replay resets after source open and
must add zero `ReadAt` calls/bytes and zero compressed Pictures overlap. A
separate selected-media replay must cover exactly the selected member's full
compressed range and reports bytes outside Pictures independently. This direct
typed replay is evidence for source laziness; production routing tests bind it
to the unified root filesystem handoff. Eager records intentionally omit source
replay vectors.

## Claims deliberately not made

This change makes no claim about latency, throughput, tail latency, physical
disk I/O, decompression volume, allocations, peak memory/RSS, cache behavior,
or release performance. Range overlap is logical positional-source evidence
for the generated fixture. A release-build, CPU-pinned, balanced ABBA capture
with retained raw samples and complete correctness gates is still required
before any performance claim.

## Verification

The focused harness test covers selector parsing/dispatch, default exclusion,
deterministic corpus reuse, root eager/source semantic and metadata parity,
source-file archive/hash parity, zero additional source query reads, exact
selected-media range coverage, and deterministic replay vectors. The opt-in
smoke command is:

```text
CARGO_INCREMENTAL=0 cargo run --manifest-path tools/perf-baseline/Cargo.toml --offline -- \
  --case odp_file_eager_open,odp_file_source_open,\
odp_file_eager_selected_slide,odp_file_source_selected_slide \
  --samples 2 --warmup 1 --json target/perf/odp-unified-root.json
```
