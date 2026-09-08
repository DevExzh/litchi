# 0469: borrowed SpreadsheetML compaction events

This bundle measures removal of temporary event ownership in the complete
changed-XML compaction pass. Production candidate `7cc58fc1b` preserves the
existing parser, normalization, whitespace and error paths. See
[source review](source-review.md) and [frozen protocol](protocol.json).

The six-row ABBA probe uses 100 retained samples and five warmups per row:
one-cell and one-percent ordinary commit/save, each at tiny, medium and
dense-wide shapes. This is descriptive diagnostic evidence, below the existing
500-sample registered latency-claim minimum. It creates no registry claim.
The full guard uses the checked default 201-row matrix at 15 samples/three
warmups. Heaptrack uses dense one-percent commit/save at five samples/one
warmup and covers the entire process. Its elapsed time and instrumented RSS
are excluded from comparisons with normal runs.

Both roles use Rust 1.98.1, release debug level 1, forced frame pointers and
unwind tables, CPU 2 and one worker. The authenticated control binary from
0468 was reused after checking its exact hash and restoring its compiled-in
manifest directory at clean revision `933cb6b80`. The candidate is built from
the same absolute directory, `/tmp/litchi-goal-0468/profile-tree`. Before each
capture that directory is switched to the lane's exact clean revision. Source
inventories bind 6,992 files and two compile-time fixtures; only `compact.rs`
differs. Binary and build receipts retain the historical identities.

All heavy commands run serially under `/tmp/litchi-goal-0469/cpu.lock`.
`capture.py` records exact argv, environment, timestamps, clean revision,
binary identity and raw artifact hashes. `export_heap.py` runs only after
captures. Raw reports, samples, corpus catalogs, process resource logs,
Heaptrack streams and exported totals are retained.

Reproduce measurements in a fresh evidence directory using the bound revisions,
source inventories, fixtures and exact build/capture commands in `build.json`
and each lane's `receipt.json`. Keep the same absolute build path for both
roles, build with the recorded environment, and authenticate freshly built
binaries before capture. Run A1, B1, B2, A2 in that order; do not mix builds,
tests or postprocessing with captures. New runs need their own receipts and
bindings rather than overwriting the historical evidence.

`analyze.py` delegates all latency statistics and guard comparison to
`tools/perf_abba_summary.py` and `tools/perf_compare.py`; the local
`report-policy.json` freezes the guard policy. `verify.py --live` checks the
retained evidence and live binaries before cleanup. Flagless `verify.py`
checks the sealed bundle and requires the temporary binaries/build tree to be
absent. Portable replay needs this bundle, the two canonical Python tools,
and the referenced adjacent 0468 binding, build and source-binding records. It does not rebuild
Rust or require the temporary binaries.

Validation receipts cover the complete XLSX test suite, scoped formatting,
workspace all-feature checking, warning-denied XLSX Clippy and rustdoc, and
crate boundaries. No new native Office, fuzz-campaign, physical-cold,
remote/range or worker-scaling claim is made. The broader non-iWork goal remains
open.

The [change record](../../changes/0469-xlsx-borrowed-compaction-events.md)
explains the retention decision and all limits. `summary.json` contains the
six-row probe, full guard and Heaptrack totals. `review-summary.json` retains
the targeted seven-row ABBA rejection, all descriptive follow-up observations,
and every full-guard mean/p50/p95/p99 flag above five percent. Run `review.py`
to replay that evidence; `verify.py` also recomputes it. Optional empty source
vectors are projected only in comparison copies, with equal-path audit.
Historical helper failures and corrections are retained under `validation/`.

All 11 focused evidence tests, live replay, fresh-copy portable replay and the
10-claim strict registry audit pass. Owned measurement binaries, the restored
worktree and portable scratch copy were removed; shared Cargo caches remain.
