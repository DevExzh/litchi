# Change 0471: reject early rewrite-buffer release after memory measurement

`performance_claim: none; rejected lifetime experiment`

`claim_authorized: false`

The candidate `aea5b744f` explicitly drops the obsolete pre-compaction worksheet
vector immediately after successful compaction, before optional grid parsing.
An independent source review establishes that this preserves output, validation
order, typed errors, no-op handling, publication and the bounded Store handoff.
It shortens an owned buffer lifetime; it does not remove parsing or allocation.

The experiment does not demonstrate the required practical memory benefit.
Both whole-process Heaptrack runs report 32,012,510 allocation calls and the
same rounded peak heap display, `104.38M`. Temporary allocation counts are
6,693,550 / 6,693,553. These include generation, expected output, warmups,
verification and teardown; they are not operation-local counts, exact peak
bytes, or proof of an unchanged peak within every individual commit phase.
Instrumented timing and RSS are excluded from normal comparisons.

Normal RSS in A1/B1/B2/A2 order is 117,576 / 118,332 / 118,216 / 116,980 KiB.
The candidates have no reduction in either matched pair. Releasing a vector
can remove local overlap without lowering process high water or returning
allocator pages to the OS. Neither that source-level rationale nor a possible
earlier peak is sufficient evidence to keep the change under the frozen gate.
The production candidate is therefore rejected; its commit remains available
for reproduction and an explicit revert restores the prior source.

Both roles use a clean checkout at `/tmp/litchi-goal-0468/profile-tree`, Rust
1.98.1, release debug level 1, forced frame pointers/unwind tables, CPU 2 and
one worker. The authenticated control reuses the 0470 candidate `f0ab67b55`.
Both inventories contain 6,993 source files and the same two compile fixtures;
exactly the transaction source differs. Seven-row ABBA has 100 samples/five
warmups: one-cell and one-percent commit/save on tiny, medium and dense-wide
worksheets, plus payload-heavy PPT creation. A 201-row default guard uses 15/3,
and Heaptrack uses dense one-percent commit/save at 5/1. These are diagnostic
measurements below the existing 500-sample registered-latency requirement.

The full non-iWork performance goal remains open. The next snapshot candidates
are documented in the [source review](../results/change-0471/source-review.md)
and [next-work analysis](../results/change-0471/next-work.md). Full parser/layout
fusion must solve MCE source identity, error order, no-op overhead and memory
overlap before implementation. No global cache, larger Store handoff, streaming,
native-producer, cold/range-source or scaling claim follows from this experiment.


## Diagnostic latency observations

Median elapsed time, milliseconds; these are not qualified speedup claims.

| Case | Shape | A1 | B1 | B2 | A2 |
| --- | --- | ---: | ---: | ---: | ---: |
| ppt_fresh_write_to | payload-heavy | 4.494354 | 4.949759 | 4.863784 | 4.910495 |
| xlsx_one_cell_commit_save | dense-wide | 170.005178 | 166.009698 | 166.253876 | 168.960067 |
| xlsx_one_cell_commit_save | medium | 2.421885 | 2.391836 | 2.382792 | 2.447017 |
| xlsx_one_cell_commit_save | tiny | 0.218071 | 0.217181 | 0.216411 | 0.217866 |
| xlsx_one_percent_commit_save | dense-wide | 335.333335 | 339.954140 | 344.306132 | 341.119081 |
| xlsx_one_percent_commit_save | medium | 9.557032 | 9.538091 | 9.438257 | 9.664749 |
| xlsx_one_percent_commit_save | tiny | 0.393632 | 0.392467 | 0.389647 | 0.392822 |

PPT control medians drift from 4.494354 to 4.910495 ms, so its first adverse
candidate pair is not a stable candidate-only effect. XLSX dense one-percent
medians also vary. The complete summary retains mean, p50, p95, p99, identity
checks and drift results for every row. These latency observations do not
override the unmet memory retention gate.

The full guard compares 1,205 metrics over 201 rows and retains 61 latency
policy flags, with no non-latency flags. An explicit five-percent trigger finds
114 mean/p50/p95/p99 cells across 44 rows; nine rows exceed it in all four
statistics. Those short-run observations remain visible without claiming a
uniformly regression-free binary. Full-guard normal RSS is 152,440 / 149,800
KiB; this single small reduction does not overcome the unchanged heap peak
and the opposite direction in both normal ABBA pairs.

All six correctness gates pass, including 1,257 XLSX tests across 59 targets.
The explicit revert is `b8b0ae445`; the XLSX tree is byte-identical to control
`f0ab67b55`. These are existing source semantics restored after a measured
rejection, not an untested alternative implementation.

See the [evidence bundle](../results/change-0471/README.md) and
[machine-readable summary](../results/change-0471/summary.json).
