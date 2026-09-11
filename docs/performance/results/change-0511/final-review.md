# 0511 bounded final evidence review

The reviewed candidate is the sector-batched FAT decode in `crates/litchi-cfb/src/file.rs`. It keeps the existing fallible reservation for the complete FAT table, sector reads, marker checks, and publication boundary, then extends the table from complete four-byte chunks with `u32::from_le_bytes`. The source manifest and candidate diff bind to that single Rust file. The earlier direct-push variant is retained only as a rejected diagnostic: its scoped five-open profile increased instruction references by 2.21%.

There is no source or completed-evidence correctness blocker in the bounded review. All final test/gate receipts now pass and are bound by the complete replay. The evidence supports a bounded OLE2 result; it does not support a package-wide performance, memory, or compatibility claim, and ODF is outside this batch.

## Checks performed

Read-only native validation passed for all eight XLS/CFB report files, report identities and catalog bindings, raw elapsed vectors, source and binary identities, ABBA ordering, the 22 paired rows, and the negative short-vector check. The native preview contains 44,000 samples: nine XLS cases and two incompressible CFB guard cases, two repeats, four children, and 1,000 measured samples per child after 20 warmups. No native adverse timing, drift, or whole-child RSS flag is set.

The allocator reports are now present and pass the same identity, corpus, operation, and serial-ABBA checks: nine XLS rows, four children, 30 samples per row, and three warmups (1,080 measured samples). All four captures have identical operation allocation vectors, including bytes, call counts, reallocations, and failures. The incremental region peak is also unchanged in each matched row. These are allocator deltas within the existing operation boundary; they are not process-peak memory measurements. The CFB guard reports allocation metrics as unavailable, which is preserved as unavailable rather than interpreted as zero.

The hardware validator passed its 4,000 samples and receipt/artifact bindings. The supplemental profile and rejected-push checks pass. The complete `verify.py` replay now also validates `gates.json`: 4,372 test executions pass with 28 ignored across the OLE2 and CFB feature checks; this includes repeated CFB coverage. Formatting, strict clippy/rustdoc, boundaries, strict claims and report classification pass.

The plan's 5% threshold is a review trigger for latency, throughput, and RSS observations, not an automatic rejection rule. The verifier preserves the flags and the admission record must carry the explicit disposition. Here the native flags are clear; the mixed owned-source rows and positive hardware-cycle changes remain part of the disposition even though they are below that threshold.

## Native result scope

The 18 XLS rows cover the three lifecycle shapes and three selected operations over the nine fixed XLS cases. Eager semantic rows have p50 changes from −14.87% to −17.25% across repeats. Source-backed rows are smaller, from −0.75% to −4.54%. Owned-source rows are mixed, from −3.51% to +1.35%, so the eager result must not be generalized to every source shape.

The two CFB guard shapes are measured separately. The tiny guard has p50 changes from −8.85% to −11.99%; the few-large guard has changes from −45.78% to −46.30%. These are fixed incompressible guard corpora and selected `OleFile::open` rows, with their documented timer and oracle boundaries. They are evidence for this corpus and operation, not a claim about all OLE2 files.

## Profile and hardware interpretation

The final source-open callgrind profile has exactly five `from_read_at_with_limits` calls in each binary and reports 15,127,968 to 13,973,601 scoped instruction references (−7.63%). The exclusive FAT-loading diagnostic changes from 1,407,020 to 250,820 references (−82.17%). These are simulated instruction references for the five-call profile and do not replace native elapsed or hardware measurements.

The supplemental legacy CFB guard profile changes from 22,681,105 to 13,081,195 runner references (−42.33%) over five opens. Its runner includes timers, file-size oracles, drops, and result construction, so this is attribution evidence that is consistent with the larger CFB result, not an operation-local or causal speedup measurement.

The hardware captures cover 4,000 whole-child samples, including fixture creation, input clones, open/query, oracles, drops, and JSON reporting. Instructions change by −4.22% and −4.14% for the two repeats, while cycles change by +0.46% and +3.64%. Event groups ran for 100% of the measured runtime. Instrumented timing is excluded, cache events were not collected, and these results do not establish an operation-local speedup.

## Final disposition

The earlier documentation inconsistencies are resolved: `cfb-source-review.md`
and `fat-reservation-review.md` explicitly preserve the rejected initial
experiment, while `sector-batch-review.md` proves the final extension. The
scope examples use the actual `guard-r1-*` artifact names. `admission.json`
is labeled historical; `acceptance.json` records final retention.

Root's final source-bound gates and complete replay passed after the read-only
review. No source, completed-evidence or gate blocker remains. Retain this
bounded OLE2 result with the unchanged allocation vectors, mixed OwnedSource
rows, positive hardware-cycle changes, rejected diagnostic and all scopes
above explicit. Cleanup and the evidence seal are recorded separately.
