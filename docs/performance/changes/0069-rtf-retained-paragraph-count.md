# Change 0069: RTF retained paragraph count

Date: 2026-08-12

Production base: `6604dd0215227086a7a0ea4a2fbfd4d3fd77519a`

Status: accepted

## Decision and scope

Retain exact visible main-story paragraph cardinality in the private RTF
parser/model handoff and return it directly from the public immutable
`Document::paragraph_count` facade. The parser counts only paragraph breaks it
already admitted as visible, non-table, non-deletion root-body boundaries. It
also retains the position immediately after the last admitted paragraph break
so final unterminated visible text contributes exactly one trailing paragraph.

The value is finalized only after the existing complete group parse, table and
range finalization, and boundary validation succeed. It is then carried through
`ParsedDocument` into the immutable `RtfDocument`. The facade no longer needs a
`OnceLock<usize>` whose cold initializer traversed and constructed every lazy
paragraph view. No source field, persisted metadata, public type, dependency,
cache, lock, unsafe code, archive abstraction, or edit/publication contract was
added.

## Semantics and boundedness

The retained count is derived from parser-owned state, never trusted from the
input. Increment and final trailing-paragraph arithmetic are checked. The
existing finite source, token, block, boundary, text and allocation limits
remain the bound.

Differential tests require the retained public count to equal complete lazy
enumeration for:

- empty, only-`\par`, consecutive-empty, terminal-break and unterminated-tail
  stories;
- inline `\line`, decoded U+000A, fragmented formatting and sparse `nth`
  traversal;
- hidden header, inert unknown destination, deleted text and table content;
  and
- first, middle, last, exact-end and `usize::MAX` selection/exhaustion states.

The existing transport/producer matrix separately verifies the same public
count through plain UTF-8, raw CP-1252, LZFu and the content-addressed
LibreOffice watermark fixture. Full parse/open, enumeration, text, exact
stream/no-op save, changed edit/save, patch/inverse and readback behavior remain
unchanged.

## Harness and protocol

Two opt-in cases were added without changing the 36-case/198-record default
matrix:

- `rtf_semantic_paragraph_count` parses outside the timer and times one cold
  public count query; and
- `rtf_semantic_collect_paragraphs` collects all lazy paragraph views as an
  allocation/layout guard separate from the historical traversal case.

The harness now exposes 130 selectable case names. Its all-variant tiny RTF
smoke has 33 rows across nine names; the tiny-plus-large scheduled matrix has
58 rows.

The unchanged large plain corpus contains 10,000 paragraphs, 10,000 retained
blocks, 499,999 visible text bytes and 540,051 source bytes. Its SHA-256 is
`957645f9109433d8dc25a66e384a496b19a97ed5ff4fab4bb981f8cda3c6e02e`.

The frozen control binary is
`360e98a4209c570949fabaff4d1438a6accc8cbcdea44217e739fa85d993259d`;
the retained candidate is
`383771bcf213ee05fa2ef02036aa2977b91475ddc62a3cd363ef2f292a028dc5`.
Both use the identical release harness. On CPU 2 the retained sequence was
before A, after A, after B, before B. Headline and collection legs used 100
warmups and 1,000 samples, yielding 2,000 samples per state. Seven ordinary
large-plain guards used 50 warmups and 500 samples per leg. Raw samples,
profiles, counters and memory are indexed in the
[`measurement summary`](../results/rtf-retained-paragraph-count-summary.json).

## Results

| Metric | Before | After | Delta |
|---|---:|---:|---:|
| pooled samples | 2,000 | 2,000 | — |
| p50 | 28.898 us | 0.020 us | **-99.93%** |
| mean | 30.302 us | 0.027 us | **-99.91%** |
| mean 95% interval | 29.732-30.873 us | 0.026-0.028 us | disjoint |
| p95 | 36.469 us | 0.050 us | **-99.86%** |
| p99 | 46.532 us | 0.110 us | **-99.76%** |

Both ordered legs improve by more than 99.8%. Candidate p50 is 20 ns in both
legs. Candidate between-leg mean drift is 3.3% (0.875 ns absolute), within the
5% policy despite the query being at timer-resolution scale.

Complete paragraph collection is a clean guard: p50 improves 1.61%, mean
improves 1.69%, p95 improves 2.86%, and p99 improves 4.26%. This separately proves
that the facade/model layout change does not make borrowed view collection
slower.

## Ordinary RTF guards

| Case | p50 | Mean | p95 | p99 |
|---|---:|---:|---:|---:|
| open | +2.04% | +2.79% | +8.91% | -22.06% |
| paragraph traversal | -0.24% | -1.17% | -1.25% | +2.06% |
| middle paragraph | +0.90% | +0.31% | -4.49% | +10.09% |
| first full text | -3.33% | +1.22% | +3.36% | +10.53% |
| exact stream save | -3.07% | -2.50% | -4.66% | -6.60% |
| exact no-op edit/save | -10.81% | -9.93% | -12.00% | -1.49% |
| one edit/save | +0.90% | +1.37% | +3.40% | +9.83% |

Every guard regression in p50 and mean is inside the 5% policy; open p95
remains inside the 10% tail gate. The noisy one-paragraph/full-text p99
movements are disclosed and are not used to claim a benefit. The parser adds
two transient scalar counters and one checked increment per admitted `\par`;
the 2.79% open mean movement remains inside policy. The no-op improvement is
treated only as a clean guard.

## Allocation, memory and CPU evidence

Heaptrack over 100 complete large samples reports allocation calls exactly
flat at 4,087,128 and peak heap flat at 14.30 MiB. Temporary allocations move
by one call (1,020,414 to 1,020,415); Heaptrack-inclusive RSS moves 25.96 to
26.04 MiB (+0.31%); the same 544 bytes of profiler/runtime leakage remain.
Uninstrumented GNU Time reports 30,976/30,848 KiB before and 30,976/30,976 KiB
after, so maximum observed RSS is exactly flat.

Three process-wide `perf stat` repeats include deterministic corpus creation,
1,010 complete parses and complete verification. Instructions fall 1.19%,
branches 0.91%, cache references 3.86%, and cache misses 5.42%; task clock,
cycles and branch misses move +1.43%, +1.67% and +3.09%, respectively. CPU
migrations are zero. These broad counters dilute the 29-microsecond scoped
query removal.

Matched 5,000-sample `perf record` profiles contain no lost samples. Exclusive
`Paragraphs::next` attribution falls from 0.99%/146 samples to 0.23%/35
samples. The residual after samples are from the complete verification pass,
which still enumerates every paragraph outside timing.

## Rejected exact-size extension

An intermediate candidate also added retained cardinality to every `Story` and
made `Paragraphs` exact-sized. That enlarged each collected borrowed paragraph
view because `Paragraph` carries a copy of `Story`. Its first matched large
collection leg moved 60.314 to 62.907 us p50 (+4.30%). Candidate binary
`04f4e8b811e41cd6a0d5ecdeb34a858829702c7f7912858c80296bb51ac4c341`
was rejected, the iterator/layout changes were fully removed, and all retained
measurements were rerun with candidate
`383771bcf213ee05fa2ef02036aa2977b91475ddc62a3cd363ef2f292a028dc5`.
The separate collection benchmark remains so future cardinality work cannot
hide the same tradeoff.

## Validation and limitations

Passed on the final source:

- the complete all-feature RTF library, integration and doc-test suite, plus
  the expanded count/enumeration structural differential;
- all 36 performance-harness tests and the 33-row plain/CP-1252/LZFu/watermark
  semantic smoke;
- warning-denied all-target/all-feature RTF Clippy and warning-denied RTF
  rustdoc;
- warning-denied performance-harness all-target Clippy;
- warning-denied ODF-common all-target/all-feature Clippy and rustdoc,
  revalidating the deprecation cleanup from `1194fbc7f`; and
- formatter, JSON, whitespace and final-diff checks.

This is an already-open generated-story scalar query. It does not make parse,
paragraph enumeration, property resolution, media extraction or edit/save
publication O(1), and it does not change the legacy `RtfDocument::paragraphs`
materializing API. Broader formatted/media, malformed/security and real-
producer matrices remain open. OLE2, OOXML, ODF production code and every
iWork/IWA crate are unchanged by this batch.
