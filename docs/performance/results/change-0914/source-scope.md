# 0914 source and timer scope

The DOCX section-layout selector uses the fixed generated corpus
`litchi-docx-section-layout-source-v1`. It contains 256 direct body
paragraphs plus one paragraph inside a table cell, so the main-story inventory
contains 257 paragraphs. The two paragraph-owned `w:sectPr` elements are at
direct XML paragraph indices 63 and 128. Because the table-cell paragraph is
counted by the semantic inventory, their public positions are 64 and 129. The
timed edit selects semantic position 64; it does not select XML index 63
directly. The table-cell `w:sectPr` is a sentinel and must remain excluded
from section ownership.

## Timed source path

Each iteration starts the timer before cloning the generated archive into an
`InstrumentedSource` and reserving the bounded sequential sink. The timer then
covers, as separate phase intervals, source-backed package open, the section
layout query, the edit that changes the selected section's columns from two to
three, commit, and sequential publication. `elapsed_ns` is the checked sum of
those six phase intervals. The source-backed query reads the main-document
payload and computes the source artifact fingerprint. Publication performs its
source checks, overlay write, and output fingerprinting within its publication
interval.

The timer stops when publication returns. Sink destruction, the enclosing
iteration's remaining teardown, and all post-publication verification are
outside the timer. In particular, semantic reopen, section and relationship
checks, patch forward/inverse/stale checks, inverse publication, refusal
controls, output digesting, and source immutability checks do not contribute to
`elapsed_ns`.

The original instrumented source is not read by those post-timer checks. Its
read counters are snapshotted after the checks, but they represent reads from
the timed open, query, and publication path (plus any source access made by a
timed phase). The output
reopen uses a separate source constructed from the emitted bytes. The counters
therefore do not provide an untimed-replay correction or a split between the
query's fingerprint work and its other parsing work.

## Evidence boundary

The source read calls, returned bytes, ordinary payload reads, cache counters,
and bounded sink counters are logical harness observations. They are not
physical device I/O, decompressed-byte, copied-byte, allocator, peak-live,
RSS, or syscall measurements. The fixed generated archive and deterministic
media establish reproducible input identity; they do not establish producer
diversity, cold-cache behavior, remote/range behavior, or a general DOCX
performance claim.

The repeated complete-artifact fingerprints visible in the query and
publication path are a measurable mechanism hypothesis for a later scoped
profile. This packet does not infer their cost from source read counters or
phase medians, and it does not authorize removing or weakening the exact-source
and stale/conflict checks.
