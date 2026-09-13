# OLE2 physical-sector marker accounting

Status: candidate rejected by mandatory native gates; exact baseline source restored. All eight final quality commands passed; matched evidence and the rejection decision are retained in change-0555.

This campaign starts at `d3c62f19a` and tests a private CFB physical-role representation. During existing FAT decoding, a non-free marker for an unclaimed physical sector becomes `Pending` in the already allocated role vector. Later successful ownership claims consume that state. The final validation scans roles in order and reads the FAT only to report an offending marker. It remains an O(P) scan; the hypothesis removes paired FAT lookups, not all validation work.

The [source review](results/change-0555/candidate-review.md) proves first-offender and short-FAT ordering from the pending invariant. FAT padding remains ignored, existing FAT/DIFAT checks retain their phase, and failed claims retain their exact errors. No bitset, cursor field, additional allocation, public API or directory/name change is introduced. Extra FAT-decoding and claim work is measured within the complete XLS constructor.

The [prospective plan](results/change-0555/plan.json) retains nine XLS workflows and three CFB shapes in two native ABBA repeats with 1,000 samples and 20 warmups. Allocation captures use 30 samples and three warmups. Profiles retain five positive timed samples for each of four constructor jobs in both repeats. Four primary XLS p50 rows must improve at least 3% in each repeat; other native controls, allocation metrics and group-process RSS must remain within the 5% ceiling. The many-small positive requirement from the different 0554 name-handoff candidate becomes a CFB control here, as recorded before captures.

The corrected candidate passed 313 CFB tests in four groups, with one existing ignored example, plus warning-denied Clippy and formatting. The initial test compilation failure and its test-only qualification correction are retained separately. Baseline profile parsing and binary custody preflight pass after a preserved metrics-consumer tuple/set correction. Neither correction changed a performance gate or discarded a measurement.

OLE2 and OOXML remain the priority. ODF is deferred until that goal completes; iWork is excluded. The independent [OOXML follow-up audit](results/change-0555/ooxml-next-opportunity.md) identifies a checked linear provenance merge as a subsequent source-backed XLSX opportunity, with fresh allocation attribution still required.

## Matched results

All native, allocator and profile captures completed in both repeats. All eight primary XLS median checks failed:

| Primary XLS workflow | Median change R1 | Median change R2 |
| --- | ---: | ---: |
| Source-backed open | +8.66% | +14.97% |
| Source-backed open + one cell | +16.16% | +14.51% |
| Owned-source open | +23.23% | +16.80% |
| Owned-source open + one cell | +23.49% | +18.58% |

The required direction was at least 3% improvement in each row. All eight primary means also failed their regression ceiling. Eight of twenty other XLS controls and four of twelve CFB controls failed. Few-large CFB p50 regressed 44.81% and 31.49%; tiny and many-small stayed within their 5% ceilings. All 72 allocation metrics were exactly unchanged, and all four group-process RSS comparisons passed. The review retains 138 adverse comparisons and 49 same-build drift rows. No selective timing rerun or gate relaxation was used.

The positively attributed XLS profile rows show the work transfer clearly. Across five timed dumps per repeat, final physical-scan self instructions fell from 1,991,710 to 829,930, while FAT-loading self instructions rose from 250,505 to 3,728,720. Collector self instructions remained 5,601,140. Total selected-owner instructions increased from 11,317,486 to 13,634,497 in repeat 1 and from 11,320,350 to 13,639,562 in repeat 2. These exclusive rows are kept separate from overlapping inclusive costs. The final scan became cheaper, but the combined constructor became more expensive.

The normal native measurements determine the failed latency gates; Callgrind instruction counts explain the measured work direction and are not substituted for elapsed time. The pending-role implementation is not retained in production.

## Validation and next work

The final restored-baseline test run passed 4,228 tests in 154 groups, with 27 existing ignored tests/examples. The candidate's three new tests remain in its retained witness; no candidate Rust change is left live. Workspace/features, Clippy, rustdoc, formatting, boundary and claim checks all passed, completing the eight-command final quality matrix.

The existing CFB fuzz target is retained, but `cargo-fuzz` and a nightly toolchain are unavailable in this environment. No fuzz or sanitizer result is claimed. The candidate was rejected on both native and complete-owner instruction gates, so this availability observation is not used to authorize production adoption.

The physical-accounting experiment is closed on its measured cost: scalar bookkeeping during FAT materialization outweighed the saved final FAT loads. The next independent OOXML opportunity is the checked linear provenance merge identified above; it requires fresh source-bound measurements, preservation checks and operation-local allocation evidence.
