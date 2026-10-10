# XLSX trait read boundaries and capability publication

This checkpoint fixes concrete correctness defects found during record 0064.
The complete unified operations goal remains active.

## Behavior

The low-level worksheet trait uses one-based cell coordinates. Its `cell` and
`cell_value` methods now refuse zero rows or columns with typed `OutOfRange`
instead of silently reading A1. The zero-based `row` method validates its index
before checking the sheet extent, so an empty sheet cannot hide an invalid row.

Fallible workbook-trait selection (active, named and indexed worksheets) now
checks source freshness before and after catalog access. Worksheet iteration
checks each pull, reports a typed source error once, and then stays exhausted.
Missing-name diagnostics no longer echo the caller's worksheet name.

The unified node enum deliberately retains inline format nodes, including in
XLSX-only builds. A documented lint exception avoids imposing a heap allocation
on every spreadsheet query node solely to balance feature-dependent enum sizes.
No allocation or throughput acceptance is claimed.

Regenerated `officecli-capability-delta.md` from the current capability goldens
and the existing base checkout's OfficeCLI schemas. The report now reflects the
PPTX adapter and retains explicit missing and extended-gap classifications.

## Remaining compatibility limitation

The old `litchi_core::sheet` metadata methods return infallible values, and its
worksheet dimension implementation can still suppress read errors. They remain
a temporary evaluator compatibility seam, not the unified read path. Closing
this requires migrating evaluator consumers to fallible typed metadata; this
checkpoint does not claim those methods detect source changes. The unified
read adapters use typed leaf methods instead of these boxed trait projections.

## Verification

Offline `cargo test -p litchi-ops --features xlsx,test-support` passed 468 tests
and doctests. Coverage includes both eager/source coordinates, empty-sheet row
bounds, all fallible catalog selectors after a source change, cached worksheet
payload reads and pre-existing worksheet iterator failure/exhaustion.

Clippy with the same feature profile and all targets passed with warnings denied.
The deterministic OfficeCLI report check and its 10 Python tests passed.

## Commit and cleanup

The rewritten local base remains an ancestor of HEAD. `git add -A` again failed
because the linked worktree index is read-only, so no commit was created. Both
required staged checks were silent and the changed/untracked artifact audit
found no prohibited evidence or binaries. Removed generated Cargo output,
untracked generated lockfiles, Python caches and checkpoint scratch files.
No temporary worktrees were created.
