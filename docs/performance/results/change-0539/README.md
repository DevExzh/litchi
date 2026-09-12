# 0539 transient XLSX attribute ownership: rejected

The [frozen protocol](protocol.md) tests the 0537 `Cow` draft with the 0538
planning allocator enabler. Both builds include the same stronger attribute
test. Baseline and candidate source manifests differ only in the raw worksheet
codec; the final codec is restored byte-for-byte to baseline. The test remains.

The [native comparison](comparison.json) validates 36 fresh child reports and
2,440 measured samples, including phase sums, source/output identities,
within-child bootstrap intervals, every adverse metric and same-build drift.
Only one of four primary rows passes the frozen workflow p50/mean and planning
p50 gates. The [decision](decision.json) rejects the production candidate.

The separate [allocation analysis](allocation-analysis.json) validates planning,
commit and publication counters for eight reports with 30 samples each. Planning
calls fall 13.5615% on medium and 13.7346% on dense-sparse in both repeats, but
planning bytes fall only 0.2639% and 0.4617%. Reallocation and incremental-peak
vectors are identical in all 12 phase pairs. These are operation-local global
System allocator observations, not physical RSS or a native latency comparison.
The legacy allocation field inside `comparison.json` is commit-only; the
separate allocation analysis supplies the planning gate and all three phases.

[Source review](source-review.md) documents the lifetime and semantic boundary.
[Adverse review](adverse-review.md) retains individual dispositions for all 28
matched and 69 same-build flags. None is dismissed to rehabilitate the failed
candidate. Conditional profiles, hardware and eager-read guards were skipped
after native rejection, as planned; no corresponding measurements are invented.

`baseline_phase.py` and `candidate_phase.py` execute the serial ABBA protocol
through the source-bound `run.py`. The latter retains both baseline executables
while the candidate checkout is active. `final_quality.py` runs the final checks
after restoration. Every child command has stdout/stderr and a receipt binding
its plan, source manifest and executable where applicable. The baseline,
candidate and restored XLSX suites each execute 1,292 passing tests (3,876 total).
Final feature, warning-denied lint/doc, formatting and boundary checks are
recorded separately in the final-quality receipts.

`verify.py` replays retained evidence read-only, including source restoration,
receipt inventory, temporal ordering and the final seal. The owned target and
test temporary files are removed after verification; `cleanup.json` preserves
their absence. Scripts refuse sealed/occupied capture outputs; reproduce raw
commands from receipts with fresh paths rather than overwriting this bundle.

See [next priority](next-priority.md) for the next bounded OLE2/OOXML target.
The rejected borrowing candidate must not be revived by relaxing its frozen
gates or treating allocation count as a workflow speedup. ODF remains deferred
until the OLE2/OOXML optimization goal completes; iWork is excluded.
