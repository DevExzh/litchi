# XLSB worksheet binary index — 2026-09-10

This batch implements bounded worksheet binary-index parsing, verification,
generation, maintenance, and source-backed cached-value lookup. Validation ran
on the source manifests retained here, with parent commit
`973a70c504950735e0f8cff7eece09364516892c`.

The public indexed handle verifies index offsets against actual worksheet
record boundaries, row headers, and first-cell anchors in 1,024-column buckets.
It retains compact metadata and managed source parts, then scans the selected
bucket for a cached scalar. Missing cells and explicit blanks remain distinct;
formula caches are returned without evaluation or number-format conversion.
Finite limits bound source bytes, records, cells, entries, and returned strings.
Preparation still materializes the compressed worksheet part and scans its
framing once. Shared strings use the existing lazy whole-table cache policy.

Writer output now derives index offsets from the finalized worksheet stream.
Typed worksheet edits maintain existing indexes transactionally. Exact no-ops
preserve bytes; unchanged row/bucket geometry also preserves unknown or
malformed index bytes without parsing. Strict indexed lookup rejects those
indexes. Offset changes patch supported index layouts, while topology changes
regenerate supported indexes and preserve formatting-only rows. Unsupported or
stale indexes that require changes cause atomic refusal. This is a bounded
implementation, not a claim that every remaining XLSB audit gap is closed.

## Validation

[Independent review](review.md) cleared the final implementation against the
normative grammar and native fixtures. The [gate receipt](gates/receipt.json)
and raw logs record Rust 1.95.0, offline locked dependencies, 847 passing tests
across 22 targets and 10 passing doctests. One library test and 11 doctests
remain ignored. Strict Clippy, warning-denied rustdoc, formatting, and diff
checks passed. Before/after hashes match all 21 frozen source/test/example
files and the retained workspace lockfile.

The release matrix contains 36 fresh-process reports: three fixtures, four
cases, and three processes per case. Each process ran three warmups and 30
timed samples, for 1,080 timed samples. Native first-sheet fixtures contain 48
and 282 cells; the generated control contains 131,072 cells. Queries select
six targets on the first native fixture and five on each other fixture.

[Root verification](root-verification.json) independently checked source and
build provenance, raw statistics, interval allocator accounting, logical
source counters, and every sample digest. A separate Python BIFF12 decoder
checked native expected values; synthetic expected values were independently
computed from the generation rule. The release binary hash was checked against
the live executable before the temporary build directory was removed.

| First worksheet | Cold indexed requested allocation | Cold materialized requested allocation |
| --- | ---: | ---: |
| `testVarious.xlsb`, 48 cells | 1,096,894 bytes | 1,227,095 bytes |
| `62815.xlsb`, 282 cells | 868,943 bytes | 1,349,201 bytes |
| Generated, 131,072 cells | 4,296,310 bytes | 71,714,427 bytes |

These are median total requested allocation bytes per timed operation,
including source open, preparation, selected queries, and teardown. On the
generated control, the median of process p50 times was 3.087 ms indexed and
18.002 ms materialized. Warm materialized lookup was faster: 0.130 microseconds
versus 0.650 microseconds indexed. Both warm cases requested zero source reads
and zero allocation bytes per timed sample on that fixture.

The API alternatives perform different preparation and validation work; these
measurements do not establish an equivalent-work or universal speedup. Timings
include allocator instrumentation on a shared host. Requested allocation and
cache counters are not RSS, and in-memory logical reads are not physical I/O.
Both cold cases materialize the compressed worksheet; the synthetic indexed
case reads slightly more source bytes. See the [profile report](performance/README.md),
[raw reports](performance/raw/final-release/),
[matrix summary](performance/raw/matrix-summary.json), and
[host environment](environment.json) for scope and provenance.

## Reproduction and integrity

Run `python3 docs/report/spec-gap-validation-evidence/xlsb-binary-index/verify-root.py`
from the repository to replay independent verification. It checks retained
build proof after cleanup and additionally hashes the original executable if
that temporary path still exists. Python assertions must remain enabled.

For fresh Rust gates, restore `gates/workspace-Cargo.lock` to the repository's
`Cargo.lock` in an isolated checkout, install Rust 1.95.0 and the locked
dependencies, then run [run-library-gates.sh](gates/run-library-gates.sh) with
`FINAL_GATE_FREEZE=1`. The script records evidence beside itself, so use a
separate checkout to preserve this capture. Release reproduction settings are
in the [profile instructions](performance/README.md); unset `PROFILE_BIN` and
`PROFILE_BUILD_MANIFEST` to rebuild after cleanup, and select a fresh output
directory. Offline execution requires dependencies already cached.

`artifact-manifest.json` records sizes and SHA-256 hashes for this evidence
directory, excluding itself. Temporary release targets, proof copies, empty
profiling stderr logs, and owned Python caches were removed after verification.
