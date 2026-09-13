# 0557 allocation instrumentation final review

Status: source-bound review passed; serial harness validation and measurement
remain required. This review is read-only with respect to the Rust sources and
makes no build, test, benchmark, profile, or capture claim.

The reviewed worktree is based on revision
`0d812aca92b608f9de8f9398fad1c379077ce708`, inspected on 2026-09-13 after
formatting. The exact reviewed bindings are:

```text
0fcec3e5c4972031b23542fd359a38401e25f32a133c77ef4c6e75facf0710d5  tools/perf-baseline/src/allocation_metrics.rs
835ba2bffa686664f0388ab5381485672e3cac5d52eab9ae3f3cc339b21a8e05  tools/perf-baseline/src/lib.rs
c947077f9517a1d39003b4234e0eeff106652920b2aa615ab48884948bb88cf8  tools/perf-baseline/tests/xlsx_planning_allocations.rs
d0117f36b425fce6f637d4256bc134b7657c556b4c102de7d390a5cc870ed80c  docs/performance/results/change-0557/alloc-design.md
00b2418b96adc351fc2371a1479b10662a0e754b600e26d0d96d36aef76dfab9  docs/performance/results/change-0557/plan.json
```

The implementation is harness-only. No production XLSX crate or public
library API is changed.

## Findings

`Region::split` snapshots the current counters and region peak while holding
the existing observer mutex, then rebases the same observer marker without
releasing the single active-region token. The staging and commit-core samples
therefore partition one continuously owned observer interval. `finish_split`
captures the final segment and the retained combined interval at one locked
boundary, so no callback can fall between those two endpoints.

The retained `commit_allocation_metrics` sample still uses the original
operation start and final endpoint. The combined region peak is the maximum of
the rebased staging peak and final commit-core peak, while each phase keeps its
own checked difference. The no-split `Region::finish` path still routes through
the original combined sample behavior.

Missing split or final peaks fail closed. `split_region` marks the process
overflow state when an active marker is missing, `ActiveRegion.split_overflowed`
keeps that failure sticky across later boundaries, and both phase outputs remain
`overflow` with numeric fields omitted. `Region::finish` consumes the active
marker even on this path and now preserves observer-invalidity precedence by
returning `unavailable` if poison or callback reentry occurs after the sticky
overflow. Checked counter addition/subtraction, checked differences, peak
regression, mutex poison, callback reentry, zero-activity segments, repeated
boundaries, missing peaks, and drop-after-error paths are covered by the added
unit cases.

`record_xlsx_optional_allocation_metric` preserves one-for-one cardinality for
every phase vector. It permits an evidence kind to remain absent for all
iterations, accepts a sample only at the current summary length, and rejects
both absent-to-present and present-to-absent transitions. Source-backed scalar
cell-value runs supply aligned plan, staging, combined, commit-core, and
publication vectors. Eager controls continue to omit allocation vectors, as
the amended plan requires; their elapsed and RSS evidence remains separate.

The runner places the split immediately after the final `edit.set` and keeps
the existing `commit_ns` clock boundaries. The staging snapshot and mutex work
are consequently inside the allocator binary's timed commit interval; the
design and plan document this overhead. The final allocation boundary follows
the existing elapsed-clock read, matching the old combined interval's endpoint
scope. The focused integration test checks measured normal/allocator field
shapes, vector alignment, counter sums, live endpoints, process peaks, and
combined region-peak reconstruction.

No source-level blocker remains for the root coordinator's serial quality
checks, frozen-stage guards, and fresh 0557 noise/capture. The earlier
`alloc-review.md` is retained as the superseded conditional review; this file
binds the formatted source after its requested sticky-invalidity correction.
