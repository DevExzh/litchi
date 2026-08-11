# Change 0066: sparse RTF paragraph selection

Date: 2026-08-12

Production base: `e013f9e4d`

Status: accepted

## Contract and implementation

`Paragraphs::nth` now scans the remaining structural-boundary slice directly
until it reaches the requested paragraph ordinal. Skipped paragraph boundaries
advance only the iterator's scalar text start. The selected paragraph still
passes through the existing `make_paragraph` path exactly once, retaining its
two block-location checks, formatting resolution, inline range and subsequent
iterator position.

The default `Iterator::nth` previously called `Paragraphs::next` for every
discarded paragraph. Selecting the middle paragraph of the deterministic
10,000-paragraph corpus therefore constructed and located 5,000 paragraph
views that were immediately discarded. The specialized path remains linear
and allocation-free, but removes those unnecessary semantic constructions.

This is a private iterator implementation. It adds no public method, index,
cache, allocation, dependency, runtime, lock, global state, unsafe code or
persisted format data. Ordinary `next`, full paragraph listing, parsing,
immutable snapshots, edits, patches, exact source publication and all limits
are unchanged.

## Correctness boundaries

The optimized path recognizes only the parser-retained `Break::Paragraph`
boundaries. Line breaks do not consume ordinals. A trailing unterminated body
remains one final paragraph, explicit empty paragraphs stay visible, and an
out-of-range selection fully exhausts and fuses the iterator. Calling `next`
after a successful `nth` resumes immediately after the selected paragraph.

Focused differential tests compare `nth(k)` with repeated `next` for first,
middle, last and out-of-range selections. They compare text, borrowed-text
eligibility, paragraph formatting and inline text/format/break signatures,
then compare the next returned paragraph. The matrix includes fragmented
formatting, consecutive empty paragraphs, structural line breaks, trailing
unterminated text and a decoded U+000A that is not a structural boundary. The
RTF fuzz target now performs a bounded sparse selection and resumes iteration
after every successful facade parse.

Raw CP-1252, LZFu and producer-watermark harness cases retain their complete
semantic verification and exact corpus identities. All parser, writer,
transaction, source-lineage and candidate-readback suites remain unchanged.

## Matched measurement

The primary fixed plain RTF corpus contains 10,000 paragraphs, 10,000 retained
blocks, 499,999 visible UTF-8 bytes and 540,051 source bytes. Its SHA-256 is
`957645f9109433d8dc25a66e384a496b19a97ed5ff4fab4bb981f8cda3c6e02e`.
The frozen control and candidate release binaries have SHA-256 values
`48ec19374e8ffd47fa6001cc25fdc451ad68cf93b4e48542474d030436d2bbba`
and
`f7ef0c89eb290f36bc91c50e33a9d25fd103fdb4021730a16fd3cb3d58af9e7a`.

Primary runs used CPU 2, 100 warmups and 1,000 samples per leg in
before/after/after/before order. Pooling both stable legs gives 2,000 samples
per state. Document parse, corpus construction and complete semantic
verification are outside the scoped timer. The environment was Rust 1.95.0,
Linux 6.8.0-101-generic, AMD EPYC 9575F and the system allocator.

| Already-open middle-paragraph query | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 12.446 us | 6.488 us | **-47.87%** |
| mean | 13.101 us | 6.819 us | **-47.95%** |
| p95 | 17.284 us | 8.742 us | **-49.42%** |
| p99 | 24.334 us | 12.457 us | **-48.81%** |

The reverse-order read guard pool used 2,000 samples per state. Open is
+0.33% p50 / -2.33% mean, complete paragraph listing -2.61%/-3.31%, full text
-0.50%/+0.03%, and exact stream save -0.78%/-1.25%. A separate 4,000-sample
no-op pool improves 7.96% p50 and 6.67% mean. A 500-sample changed edit/save
pool improves 2.82% p50 and 2.56% mean. Every central and p95 guard remains
within the 5% non-regression policy.

Heaptrack reports exactly 4,087,326 allocation calls, 1,020,512 temporary
allocations and 14.30 MiB peak heap in both states. Two matched uninstrumented
GNU Time pairs report 30,848/30,976 KiB before and 30,848/30,720 KiB after.
The query removes no allocations; it removes repeated cursor-location and
paragraph-view construction.

Process-wide `perf stat` includes corpus construction and a complete untimed
parse/verification after every sample. Instructions improve 0.38%, while
task-clock and cycles move +2.24%/+2.87%; these whole-process values are
attribution context, not the scoped latency claim. In the corresponding
process-wide sampled profile, `Paragraphs::next` falls from 1.41% to 0.99%
self time despite the unchanged full-list verification, while the new `nth`
accounts for 0.16%.

Raw distributions, the machine-readable summary, Heaptrack reports, GNU Time
records, and `perf` reports/counters are under
[`results/rtf-sparse-paragraph-nth`](../results/rtf-sparse-paragraph-nth/summary.json).

## Validation

Passed on the final source:

- all-feature `litchi-rtf` unit, integration and doctest suites;
- focused iterator equivalence plus CP-1252, LZFu and producer-watermark
  harness verification;
- updated RTF fuzz-target compilation;
- warning-denied all-target/all-feature RTF Clippy and rustdoc;
- all 35 deterministic harness tests and warning-denied harness Clippy;
- formatter and whitespace checks;
- warning-denied all-target/all-feature `litchi-odf-common` Clippy and
  rustdoc, retaining the earlier deprecation cleanup gate.

This tranche changes only RTF and performance evidence. OLE2, OOXML, ODF and
all iWork/IWA crates are unchanged.
