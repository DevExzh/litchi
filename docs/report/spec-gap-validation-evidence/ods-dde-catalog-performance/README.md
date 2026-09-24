# ODS inert DDE catalog performance

This evidence measures the existing DDE stage, commit, and metadata-edit
workflows on three bounded synthetic content.xml shapes. The read-only
baseline is captured at `ae9bf7bfa11eed50e6560be166eaa18987cda2b5`; a later
scoped candidate changes only `crates/litchi-ods/src/model/dde/transaction.rs`
for catalog/preallocation work.

Baseline and candidate harnesses are retained under their respective
directories. Measurements use three warmups, 15 measured iterations, CPU 2,
and `/usr/bin/time -v`; Work and Memory are budget counters and RSS is an OS
measurement.

## Validation

Final gates pass 772 tests across 45 targets, strict clippy, rustdoc, doctests,
and scoped formatting. The ODS crate currently has no doctests. All nine
implementation/test source hashes remain unchanged across the gate sequence.
The two added regressions exercise 1,024-link reorder/inverse under finite
memory and atomic refusal with scratch-memory release under a tighter budget.
The final public authoring replay is byte-identical to the preoptimization
artifact and passes the whole-content ODF 1.4 schema and DDE cache prose checks.
Commands, output, and source manifests are in [gates](gates/).

[Resource review](resource-review.md) verifies exact-count admission, overflow,
cancellation, mismatch fallback, unchanged validation, and allocation-error
cleanup order. No safety or preservation check was removed.

## Measured result

Many-link commit allocation requests fall from 448,348,123 to 12,352,299 bytes
(97.24% lower). Retained Memory delta stays at 1,463,839 bytes. Commit p50 is
20.39 ms before and 20.59 ms after on the shared host; this does not establish
a latency improvement. Allocation requests do not measure bytes physically
copied by the allocator. See [report](report.md) and [provenance](provenance.md).

All 36 benchmark rows were checked against raw stdout, process RSS, and exit
status. The isolated candidate passes 47 DDE tests. [Root verification](root-verification.json)
records final hashes, gate results, patch replay, and cleanup.
