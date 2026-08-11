# Change 0064: retained RTF story length handoff

Date: 2026-08-12

Production base: `761b88bdc`

Status: accepted

## Contract and implementation

The RTF parser already computes the exact total UTF-8 byte length while it
moves validated body `StyleBlock` text into the immutable document model. The
public borrowed `Story` now receives that private scalar together with the
same block and boundary slices. `Story::len` is constant-time and
`Story::is_empty` compares the retained count with zero; paragraph and inline
iterators reuse the same value.

The previous implementation recomputed the total by traversing every retained
block whenever a story or paragraph iterator needed its end position. The
large deterministic corpus has 10,000 blocks, so both listing paragraphs and
selecting the middle paragraph began with a complete unrelated block scan.

This is a private state handoff. It adds no public method, dependency, cache,
runtime, lock, global state, unsafe code or persisted format data. Text bytes,
block and boundary order, formatting, paragraph identity, immutable snapshot
sharing, exact source publication, limits and transaction behavior are
unchanged.

## Correctness boundaries

The length remains computed inside the complete parser/model conversion; it is
not trusted from source metadata. Existing parser limits bound retained text,
and complete boundary validation still proves UTF-8 positions against the
actual blocks. Full-text materialization continues to copy every retained
block, and every query iterator continues to traverse and validate its own
semantic range.

Focused tests compare `body().len()` and `body().is_empty()` with the flattened
text for empty, fragmented formatting, empty/trailing paragraphs and embedded
Unicode line-feed input. A changed paragraph snapshot proves that the new
snapshot receives its recomputed length while the original immutable snapshot
retains its original value. Raw CP-1252, LZFu and producer-watermark fixtures
also prove exact length and unchanged byte publication.

## Matched measurement

The fixed plain RTF corpus contains 10,000 paragraphs, 10,000 retained blocks,
499,999 visible UTF-8 bytes and 540,051 source bytes. Its SHA-256 is
`957645f9109433d8dc25a66e384a496b19a97ed5ff4fab4bb981f8cda3c6e02e`.
The frozen control and candidate release binaries have SHA-256 values
`3f4aea56e8c1921072afed0e2c2d21f87472eb4480babfee50e3f18354f4ef0e`
and
`748ae86dc8d8981c04d95c9228f80b798e7d8677b36bc2fdbc76d146272926db`.

Primary runs used CPU 2, 100 warmups and 1,000 samples per leg in
before/after/after/before order. Pooling both stable legs gives 2,000 samples
per state. Document parse, corpus construction and complete semantic
verification are outside the scoped timer. The environment was Rust 1.95.0,
Linux 6.8.0-101-generic, AMD EPYC 9575F and the system allocator.

| Already-open large query | Before | After | Delta |
|---|---:|---:|---:|
| list paragraphs p50 | 29.692 us | 25.225 us | **-15.04%** |
| list paragraphs mean | 31.116 us | 26.851 us | **-13.71%** |
| list paragraphs p95 | 39.624 us | 36.199 us | **-8.64%** |
| middle paragraph p50 | 18.926 us | 13.780 us | **-27.19%** |
| middle paragraph mean | 19.665 us | 14.705 us | **-25.23%** |
| middle paragraph p95 | 24.383 us | 20.858 us | **-14.46%** |

The reverse-order 2,000-sample guard pool measured open at +3.43% p50 /
+2.62% mean, full text at +3.99%/+4.15%, exact stream save at
+0.69%/+0.65%, and exact no-op edit/save at +3.56%/+2.19%. All central and
p95 guard metrics remain within the 5% policy. A separate 500-sample changed
edit/save guard improved 2.11% p50 and 2.68% mean.

Heaptrack reports exactly 4,488,027 allocation calls, 1,120,562 temporary
allocations and 14.30 MiB peak heap in both states. Two matched uninstrumented
GNU Time pairs report the identical 30,976/30,848 KiB maximum-RSS sequence in
both states. The scoped query removes no allocations; it removes only the
redundant traversal.

Process-wide `perf stat` includes corpus construction and a complete untimed
parse/verification after every sample. It therefore dilutes the query frame:
instructions move from 74.432 to 74.300 billion (-0.18%), while task-clock and
cycles move +2.19%/+2.53%. The exact `perf record` report no longer contains
the separately sampled `Story::paragraphs` length-scan symbol (0.40% before),
but the complete-process profile is retained only as attribution context, not
as the scoped latency claim.

Raw distributions, the machine-readable summary, Heaptrack reports, GNU Time
records, and `perf` reports/counters are under
[`results/rtf-story-length`](../results/rtf-story-length/summary.json).

## Validation

Passed on the final source:

- all-feature `litchi-rtf` unit, integration and doctest suites;
- focused fragmented/edit, CP-1252, LZFu and producer-watermark length tests;
- warning-denied all-target/all-feature RTF Clippy and rustdoc;
- formatter and whitespace checks;
- warning-denied all-target/all-feature `litchi-odf-common` Clippy and
  rustdoc, retaining the earlier deprecation cleanup gate.

This tranche changes only RTF and performance evidence. OLE2, OOXML, ODF and
all iWork/IWA crates are unchanged.
