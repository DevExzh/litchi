# 0884 Producer XLSX full-cell scan harness

The explicit `xlsx_producer_medium_source_full_cell_scan` selector adds a
current-source workload for scanning all cells of the generated producer-shaped
XLSX. It does not change production code, corpus generation, or the default
41-case matrix. Debug-profile semantic and report qualification passes; a fresh
release baseline is still required before pricing a production optimization.

## Workload and oracle

Each iteration opens a fresh `SourceBackedWorkbook` over owned positional input,
enumerates four worksheets, and visits their stored cells with `Rect::ALL`.
The timer includes opening, enumeration, production freshness checks, ordered
callback validation, and canonical callback hashing. Corpus/source preparation,
the harness's extra source-version check, aggregate tolerance-report validation,
endpoint observations, final digest comparison, and handle destruction are
outside the clock. Allocation observations bracket the same workload; process
observations include their documented probe overhead.

Every scan must produce 4,096 cells across four 32×32 sheets: 3,686 numbers and
410 strings. The runner pins the reviewed counts and canonical digest
`8d29b40718aa9d5b08c746f841195bba17ab8e01608b12e6093daf160c00764b`.
That runtime expectation is independent of the generator helpers. Ordered
sheet/address/type/value records detect missing, duplicate, reordered, or
changed callbacks. A dirty tolerance report or changed source version rejects
the iteration.

The unchanged archive is 22,739 bytes, with SHA-256
`55e0901c093c27d80bcd4e5fef888f00798654cdef8c0d79d130a2ce7cbe0ae8`.
Its marker census identifies four worksheets ineligible for the streaming route.
This is eligibility evidence, not an observed runtime profile. The workload is
in-memory and synthetic; no physical-cold or native Office claim follows.

## Qualification

The [final checks](results/change-0884/final-v2/) pass formatting,
all-feature/all-target compilation, four focused tests, and warning-denied
harness Clippy. Tests cover the pinned oracle, complete scan and aligned sample
vectors, rejected incomplete/duplicate/out-of-order callbacks, and explicit-only
selector registration. The retained first compile failure was an unresolved
module path in the selector test; qualifying that path fixed it.

[Normal and allocator CLI qualification](results/change-0884/qualification/)
uses debug builds, one warmup and three measured samples per mode. The
[replay](results/change-0884/qualification-review.json) confirms six complete
sample oracles, aligned report vectors, identical producer sidecars, and zero
failed allocations. Debug elapsed values are retained as raw qualification
output and are not a performance baseline or speedup estimate.

```sh
python3 -B docs/performance/results/change-0884/verify.py
```

The [summary](results/change-0884/summary.json) binds the tested source. Fresh
release-profile measurements on the committed harness are next; no deleted
0882 evidence or old timings are reused.

## Cleanup

Seven clean, completed performance worktrees were removed without force.
Three detached source snapshots were first retained with local
`archive/performance/` tags. Two historical worktrees containing uncommitted
changes or evidence were preserved, as were unrelated teams' worktrees. The
[cleanup record](results/change-0884/worktree-cleanup.json) records exact paths
and retained references. Applied agent patch files were also removed.

The broader non-iWork performance program remains open.
