# Change 0056: indexed DOC PAPX containment

Date: 2026-08-12

Status: accepted

## Decision

Use the parser-normalized ordering already present in native DOC text pieces
and PAPX runs to resolve paragraph-terminator containment with predecessor
binary searches instead of restarting two linear scans for every paragraph.

The change is private to `litchi-doc`. It adds no public API, stored index,
allocation, cache, runtime, lock, dependency, unsafe code, or persisted state.

## Problem and attribution

`body_text::source_paragraphs` calls `RevisionEditor::is_in_table_at_cp` for
every main-story paragraph terminator. The control implementation found the
containing CLX piece and then the containing PAPX run with two
`iter().find(...)` scans from the beginning. On the deterministic 512-paragraph
DOC, that is about 131,328 PAPX probes per paragraph-list operation.

An exact control profile of the full one-edit/save case attributed 5.52% of
process self cycles to `is_in_table_at_cp`: 2.84% in the timed edit target
resolution and 2.68% in the mandatory post-timer verifier. The new direct
already-open snapshot case isolates the same work. Its matched control profile
attributed 22.95% of sampled cycles to the function.

## Implementation

For each lookup, the editor now:

- partitions the ordered piece slice at the last `start <= cp` candidate and
  checks the existing half-open end boundary;
- converts CP to FC exactly as before;
- partitions the ordered PAPX slice at the last `start <= fc` candidate and
  checks the same half-open end boundary;
- retains the original missing-piece, overflow, missing-PAPX, SPRM decoding,
  strict-SPRM, and in-table result behavior.

The parser already validates and normalizes both collections into start order.
The helper does not rely on interval overlap: these containment tables are
non-overlapping at this call boundary. Empty slices and gaps still return no
match.

## Correctness and safety boundaries

Scalar/indexed differential tests cover empty and singleton slices, adjacent
and gapped intervals, exact starts and ends, and `u32::MAX` boundaries for both
pieces and PAPX runs. They also prove that the same stored object is selected,
not merely an equivalent value.

The change does not alter decoded text, paragraph membership, table filtering,
formatting, revisions, fields, pictures, comments, glossary stories, patches,
inverse operations, publication, or validation limits. Malformed gaps still
produce the same typed corruption errors. No speculative memory is retained.

## Measurement method

Base revision: `1198544231fc2b5be8fc251510d42a26a0df81f0`.

Exact release binaries, built from one identical harness:

- before: `c6a7c88c6b28f53875fca83598952c61329b16859c473bba451720a892c37e33`
- after: `316dba3dd63112de27f05c8dcaa551f148ad79fd441985ad8dd35c437550da98`

The primary measurements use five balanced pairs on CPU 11. Every leg has 50
warm-ups and 500 timed samples in before/after/after/before order, yielding
5,000 samples per state. The corpus is a deterministic 512-paragraph DOC:
97,792 archive bytes, archive SHA-256
`3d96764fe48e213b972ff5921df183dab9e8bfc8c8e751bcf3bf20190de4fec6`,
and WordDocument SHA-256
`33e6cd70a45181c28d4a3e7bfa4e7817bd82d7b2e89e39437a589243abdc38eb`.

The direct case opens one exact-source `body_text::Snapshot` before timing,
times `Snapshot::paragraphs(Projection::All)`, and verifies every paragraph
position and text outside timing. The end-to-end edit case retains its complete
candidate, patch/inverse, strict owner, and independent public readback checks.
Raw pooled samples and distributions are in
[`doc-papx-containment-primary-summary.json`](../results/doc-papx-containment-primary-summary.json).

## Latency result

| Case | Before p50 | After p50 | p50 | Mean | p95 | p99 |
|---|---:|---:|---:|---:|---:|---:|
| Already-open body snapshot paragraph list | 206.644 us | 168.142 us | **-18.63%** | **-19.04%** | **-17.98%** | **-20.81%** |
| One-paragraph edit/save | 888.602 us | 817.424 us | **-8.01%** | **-7.88%** | **-7.71%** | **-8.37%** |

Every one of the five independently pooled A+B pairs improves in both p50 and
mean for both cases.

## Guard cases

Two balanced pairs cover the unchanged ordinary DOC open/list/one/full-text
reader paths and the exact no-op transaction, with 1,200 samples per state and
case. These paths do not isolate the optimized exact-source snapshot resolver.
List/one p50 move +0.61%/+0.69%; the exact no-op p50 is -0.05%; full-text is a
sub-microsecond accessor. Open improves 10.78%, but no open-path claim is made
because this capacity-free helper is not exercised there and compiler/process
layout can move an independent microbenchmark.

The tiny direct case stays on the binary-search path and moves +0.65% p50 /
+0.38% mean, with p99 -1.49%. The 41 ns p50 movement is disclosed as the small
fixed cost of binary search on a three-paragraph corpus. Raw distributions are
in
[`doc-papx-containment-guards-summary.json`](../results/doc-papx-containment-guards-summary.json).

## Allocation and memory evidence

Matched Heaptrack runs use 20 warm-ups and 200 direct paragraph-list samples.

| Metric | Before | After | Change |
|---|---:|---:|---:|
| Process allocation calls | 1,296,614 | 1,296,614 | flat |
| Peak heap | 2.11 MiB | 2.11 MiB | flat |
| Heaptrack-inclusive RSS | 11.90 MiB | 12.00 MiB | +0.10 MiB |
| Leaked bytes | 544 | 544 | flat |

The helper allocates no memory. Full reports are
[`before`](../results/doc-papx-containment-before-heaptrack.txt) and
[`after`](../results/doc-papx-containment-after-heaptrack.txt).
Four uninstrumented GNU Time runs per state report mean maximum RSS of
30,784 KiB before and 30,816 KiB after (+0.10%, flat at process resolution).
The raw reports are stored under `results/doc-papx-containment-time-*.txt`.

## CPU evidence

Two matched before/after/after/before `perf stat` pairs time 3,000 direct
paragraph-list samples per process.

| Counter | Before, mean/process | After, mean/process | Change |
|---|---:|---:|---:|
| Cycles | 4.015 billion | 3.467 billion | **-13.66%** |
| Instructions | 21.578 billion | 15.938 billion | **-26.13%** |
| Branches | 5.974 billion | 3.614 billion | **-39.50%** |
| Branch misses | 7.959 million | 4.544 million | **-42.91%** |
| Cache misses | 5.743 million | 5.687 million | -0.97% |

Exact matched profiles reduce `is_in_table_at_cp` from 22.95% to 11.10% of
sampled cycles while total sampled cycles fall from 4.018 to 3.527 billion.
There are no lost samples; only restricted kernel-symbol warnings remain.
Reports are
[`before`](../results/doc-papx-containment-before-perf-report.txt) and
[`after`](../results/doc-papx-containment-after-perf-report.txt); raw counters
are stored under `results/doc-papx-containment-stat-*.csv`.

## Validation

Passed on the final source:

- the focused scalar/indexed containment differential tests;
- the deterministic direct-case harness test;
- complete DOC tests plus warning-denied Clippy and rustdoc;
- DOC fuzz-bin compilation;
- the 33-test performance harness and warning-denied all-target Clippy;
- formatter, JSON parsing, link-target, and whitespace checks.

The ODF deprecation fixed in commit `1194fbc7f` was rechecked with
warning-denied `litchi-odf-common` Clippy and rustdoc; both remain clean.

## Limitations and next work

This is a generated warm-memory text corpus. It adds no new CRUD, media,
formatting-edit, repair, security, real-producer, cold-source, or streaming
coverage. A broader OOXML source-backed cell transaction, ODF media-preserving
publication follow-up, or a freshly profiled RTF lexer reservation remains a
separate tranche.
