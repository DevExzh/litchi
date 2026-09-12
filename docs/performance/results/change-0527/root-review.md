# Root disposition review

The frozen pilot required every primary shape/repeat to clear 2% total p50,
2% total mean and 5% commit p50 reduction, and every allocation shape/repeat
to clear 8% fewer allocation calls. The dense repeat-2 total median/mean
reductions of 1.8214108302% and 1.8105018625% fail this contract. The candidate
is rejected despite meeting every commit and allocator gate. Thresholds,
sample counts, repeat order and the frozen driver were not changed.

Comparison SHA-256: `1c973a52ef5785a5f5e2264b2429ea8267119ee14b29513c2f9c9440d9f67eab`.
Adverse review SHA-256: `78975587e89608e17d23c9864c359ad20c49564da383b2a56b95631621e30498`.
Final source manifest SHA-256: `9af673c4c13f2abb3aeaf5c4e31df6a613de6297abc104bf472931ff7733529f`.

The candidate ran 2,440 measured native durations and 40 allocation samples.
Its successful 1,294-test preflight is preserved with the same source manifest
as the measured candidate. Three production files were restored byte-for-byte
to baseline, and the two private arena tests were removed. The only retained
Rust changes are four baseline-compatible regression tests in three files.
The final quality lane executes on this restored source independently.

I reviewed the per-shape native and allocator admission rows and the complete
flag inventory. The 47 adverse flags are 12 open, 20 reopen and 15 publication
metrics; the 71 drift flags are 35 open, 18 reopen and 18 publication metrics.
Each original flag remains individually bound to its review. There are no
matched total/commit, RSS or allocation-peak flags above 5%; that absence is
not a stability or universal memory claim. Allocation peak actually increases
slightly. No metric is dismissed as noise or given an unproven causal story.
Rejecting the production candidate prevents adoption of these measured
tradeoffs and does not erase them from the evidence.

No profile, hardware or eager capture was admitted after pilot failure. No
instruction-count gain, cold-cache result, remote/range result, producer
compatibility expansion, scaling improvement or retained production speedup
is claimed. The accepted 0525 behavior remains the runtime baseline.

The initial baseline build succeeded but copying its executable into tmpfs
failed with errno 122 before any capture. The zero-byte partial file was
removed and the exact owned scratch path was linked to a disk-backed owned
subdirectory. The recovered binary hash and original successful receipt are
recorded in storage-recovery.json. This changed storage placement, not source,
build flags, capture paths or timing sample selection. A2 retained the original
baseline binary under the independently bound candidate checkout.

The source review and patch replay verify the final restoration. All planned
quality receipts, their logs and any failures remain in the bundle. Cleanup
must follow completed jobs and pre-cleanup evidence verification; final seal
verification must follow cleanup. OLE2/OOXML remains the active priority, ODF
is deferred, iWork is excluded, and the overall goal remains active.
