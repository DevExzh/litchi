# 0105 — Reuse explicit XLSX merge-list planning

An explicit `merge` list is constant throughout one set operation. When a cell
selector matches several cells in the same worksheet, the merge planner now
expands that list once for that worksheet. Previously each selected cell
rebuilt the requested-range set and scanned the worksheet's native merge
iterator, even though intent deduplication discarded the repeated results.

The operation-local set records successfully planned worksheet positions,
including lists that are entirely exact no-ops. It retains no geometry model
and does not survive the operation. Boolean range merges still plan each
range independently; explicit range-target equality is checked before reuse.
Every selected node remains subject to the expansion limit, and companion
cell assignments still apply to every selected cell. Native merge guards,
atomic commit and source checks remain authoritative.

This removes repeated list expansion and catalog scans from this admission
path. It is not a measured latency result or full performance acceptance.
Other merge-selector costs, broader format writes and Python/CLI acceptance
remain open.

Validation:

- 914 combined DOCX/XLSX/encryption facade tests and doctests passed offline.
- All 18 focused merge tests passed, including the three new regressions for
  multiple cells on multiple worksheets, selected-node limits after planning
  reuse, and disjoint companion payloads in one native commit.
- Independent read-only review found no actionable defect in limits, range
  validation, exact no-ops or source fencing.

The rewritten base remains an ancestor of this branch, so no rebase replay
is needed against the current base.


Warning-denied Clippy passed for ops and XLSX with all targets and the combined
DOCX/XLSX/encryption features. Scoped formatting and whitespace checks passed.
The 458-path artifact audit found no oversized evidence or compiled binaries;
both mandatory staged-file checks printed nothing.

The borrowed Clippy fixture, root Cargo.lock, temporary logs and 12.3 GiB of
build outputs were removed. No temporary worktree was created.

Commit remains blocked: staging failed because the shared worktree Git metadata
is read-only (`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`).
The verified changes remain in the working tree. The full goal remains active.
