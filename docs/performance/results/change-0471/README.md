# 0471: rejected pre-compaction buffer lifetime experiment

The candidate drops the obsolete rewrite vector before parsing the independent
compacted vector. The frozen [protocol](protocol.json) requires a practical
measured reduction in peak live heap or normal RSS to retain it. Both rounded
Heaptrack peaks are `104.38M`, and normal RSS is not lower in either matched
pair. The production change is rejected; retained commits and captures permit
reproduction of the negative result. See the [change record](../../changes/0471-xlsx-rewrite-buffer-lifetime.md).

Control revision `f0ab67b55` reuses the authenticated 0470 candidate binary,
SHA-256 `6ca3ace7edabc780490251ad9329ff4cc750718873c95ecd5a81bf347d3c12f7`.
The original binding, source binding, build receipt and build logs are copied
under `prior/` and authenticated by the current binding. Candidate revision
`aea5b744f` has binary SHA-256
`24126b84b0d44ca99a3fa1fde6e18c19f94d53989fb748fd1d5eb392d782c904`.
Both inventories have 6,993 files; exactly the transaction source differs.
The same two compile-time fixtures are separately bound. All previously read
accepted ADR hashes are unchanged.

Builds use Rust 1.98.1, release debug level 1, frame pointers/unwind tables,
four Cargo jobs and no incremental compilation. Both binaries use the same
absolute checkout `/tmp/litchi-goal-0468/profile-tree`. Each capture switches
that clean tree to its bound role revision and verifies all source hashes,
fixtures and the immutable binary. Measurement and compiler work is serialized under
`/tmp/litchi-goal-0471/cpu.lock`, with benchmark CPU 2 and one worker.

`build.py`, `capture.py`, `export_heap.py` and `validate.py` retain exact commands,
environment, timestamps and hashes. Reproduction requires new authenticated
role binaries from the recorded revisions and source inventories. Run the
normal A1/B1/B2/A2 lanes, A-full/B-full, A-heap/B-heap, then exports and gates.
Do not overwrite the historical captures. The normal seven-row experiment is
100 samples/five warmups; the default 201-row guard is 15/3; dense one-percent
Heaptrack is 5/1. A future qualified latency claim would need a separate frozen
protocol meeting the unchanged 500-sample requirement.

Heaptrack records whole-process generation, expected output, warmups,
verification and teardown as well as commits. Its allocation totals and rounded
peak display are not operation-local or exact peak-byte measurements. Its
instrumented timing and RSS are excluded from normal comparisons. The complete
raw samples, control drift, regression flags, output oracles and policy remain
retained; no geometric mean can hide an individual regression.

`analyze.py` replays the canonical repository comparison tools and reports every
full-guard policy flag plus the explicit five-percent latency trigger. Only the
full guard comparison treats identical empty optional source vectors as absent;
raw reports are unchanged. `verify.py --live` checks the live binary/source
identities; flagless verification checks the sealed bundle after cleanup.
Portable replay needs this complete bundle and only
`tools/perf_abba_summary.py` and `tools/perf_compare.py` in the same relative
repository layout. It needs no Rust compiler, Git checkout or temporary binary.

No new native Office execution or fuzz campaign, exact peak-memory reduction,
latency speedup, physical-cold/range, bounded streaming or scaling claim is made.
The complete non-iWork goal remains active. The rejected lifetime change does
not resolve the larger eager-parser and snapshot-layout traversals.

All six correctness gates, nine evidence tests, the live source/binary audit,
and the strict registry audit of ten existing claims pass. A fresh portable
copy verifies after owned temporary builds and binaries are removed; that copy
is also removed. Receipts are under `validation/` and in `cleanup.json`. The
shared Cargo cache is retained and contains the rejected benchmark candidate;
future production-control comparisons must rebuild/authenticate their binary.
