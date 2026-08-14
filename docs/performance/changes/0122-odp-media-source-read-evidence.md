# Change 0122: ODP media-rich source-read evidence

Date: 2026-08-15
Status: opt-in correctness and logical-range evidence; no latency claim

## Scope

The standalone performance harness adds four matched selectors over the
existing deterministic ODP media-rich corpus:

| eager control | source-backed control |
|---|---|
| `odp_media_eager_open` | `odp_media_source_backed_open` |
| `odp_media_eager_one_slide` | `odp_media_source_backed_one_slide` |

The corpus remains the fixed 12-slide presentation with eight deterministic
2 MiB resources under `Pictures/`. The middle slide is selected by its fixed
zero-based position. All four selectors are opt-in: the selectable matrix
grows from 229 to 233 names while the default matrix remains 36 cases / 198
records.

## Timing and parity

Eager open owns a prepared byte-buffer clone before timing and times only
`Presentation::from_bytes`. Eager one-slide prepares its presentation before
timing and times only the middle-slide selector. Source timing follows the
same boundary with an uninstrumented `litchi_core::OwnedSource`; source
construction is outside the query-only interval. Warmups are excluded from
the evidence vectors. Full eager/source slide projections, canonical semantic
digests, selected-slide text, and deterministic media payload digests are
checked outside timing. In particular, the eager one-slide selector performs
its full-slide and selected-media parity checks after the timed query; those
checks do not widen the query interval.

## Untimed source replay

Each measured source sample has an independent `InstrumentedSource` replay for
the named phase. The report records exact `ReadAt` calls, bytes, coalesced
prior-range overlap, and overlap with every compressed `Pictures/*` payload
range. The `pictures_read_compressed_range_bytes` vector is the latter
compressed-range overlap; it is separate from the prior-read overlap counter
`source_read_range_overlap_bytes`. The open replay includes the source-owner
catalog reads; the selected slide replay resets after source open because the
validated content XML is retained by the source facade, so that query is
expected to add zero `ReadAt` calls and bytes. A ZIP end-of-archive scan can
physically overlap the last Pictures range; this is retained as honest
request-overlap evidence and is not interpreted as media materialization.

A second untimed replay then performs one explicit selected-media read. It
checks the returned uncompressed payload and SHA-256, requires the aggregate
overlap with all Pictures ranges to equal the selected member's complete
compressed range (`selected_media_compressed_range_bytes`), and cross-checks
the same range with a selected-only replay. Bytes outside Pictures are
reported separately. The summary keeps compressed-range byte totals distinct
from uncompressed payload byte/digest fields. This makes the media-read proof
distinct from open/query catalog overlap.

The source summary stores the exact phase and timing scope, adapter names,
selected slide/member, canonical semantic digest, uncompressed media digests,
compressed-range byte totals, source vectors, Pictures vectors, and
selected-media vectors (`selected_media_read_prior_range_overlap_bytes` and
`selected_media_read_compressed_range_overlap_bytes`). Eager records retain
the same semantic and media identity fields but intentionally omit source-read
vectors.

## Claims deliberately not made

This change makes no claim about latency, throughput, tail latency, physical
disk I/O, decompression volume, allocations, peak memory/RSS, cache behavior,
or release performance. Range overlap is a logical positional-source
observation for the generated fixture. A release-build, CPU-pinned, balanced
ABBA capture with retained raw samples and complete correctness gates is still
required before any performance claim.

## Verification

The focused harness test covers selector parsing/dispatch, default exclusion,
deterministic corpus generation, eager/source semantic parity, zero additional
query reads, exact selected-media range coverage, and deterministic replay
vectors. The opt-in smoke command is:

```text
CARGO_INCREMENTAL=0 cargo run --manifest-path tools/perf-baseline/Cargo.toml --offline -- \
  --case odp_media_eager_open,odp_media_source_backed_open,\
odp_media_eager_one_slide,odp_media_source_backed_one_slide \
  --samples 2 --warmup 1 --json target/perf/odp-media-source-read.json
```
