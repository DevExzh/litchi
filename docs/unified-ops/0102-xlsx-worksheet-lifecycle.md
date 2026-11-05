# 0102 — XLSX worksheet lifecycle

The unified API now dispatches worksheet add/remove and tab rename, visibility,
move and swap to native immutable workbook edits. Direct calls, closure edits
and atomic/best-effort batches share that implementation. Structural commands
close pending cell runs; native patches support replay and inverse. Commit
adaptation is shared with cell writes and preserves exact native no-ops.

The static sheet property registry admits `name` and `hidden` for both set and
add, and supplies the capability projection. Missing creation names default to
`Sheet{count+1}`, matching OfficeCLI. Explicit invalid names, duplicate names,
unsupported properties and invalid placements retain typed errors. Native
formula-reference rewriting, protection/dependency checks and last-visible-tab
guards remain authoritative. Removal invariant errors do not misattribute a
`hidden` property that the caller never supplied.

Scoped sheet selectors deduplicate targets, reject unsupported predicates and
enforce expansion/staging budgets before native publication. Cell/range targets
continue through their own bounded writer. Root sheet names that resemble cell
addresses are distinguished from scoped cell paths. All publication routes
retain the existing source revision fences.

Non-worksheet tab removal, OfficeCLI `ifExists`/placeholder reuse, additional
sheet properties, wider structural families and broader coalescing remain
explicit capability gaps. This checkpoint makes no structural performance
acceptance claim; the complete performance and all-format goal remain open.

Validation:

- 883 combined DOCX/XLSX/encryption tests/doctests passed.
- 419 minimal-feature tests/doctests passed.
- After the final direct-cell fast-path adjustment, 34 targeted worksheet,
  cell coalescing and target-write tests passed. These include source/eager
  lifecycle, replay/inverse, formula-reference rewriting, all placement forms,
  nonadjacent/reversed swaps, A1-like sheet names, atomic/best-effort failure
  behavior, unsupported predicates and expansion/staging rollback.
- Capability snapshots and 14 generated-delta tests passed.
- Independent read-only review found no remaining actionable source-fence,
  target-selection or native-semantics defect.

Cargo ran offline. No leaf XLSX implementation was changed in this checkpoint.
The rewritten base remains an ancestor, so no rebase replay is needed against
the current base. No temporary worktree was created.

Warning-denied Clippy passed for ops and XLSX, all targets, with
DOCX/XLSX/encryption enabled. Scoped formatting, whitespace checks and the
448-path repository artifact audit passed; both required staged-file checks
printed nothing. The borrowed LibreOffice Clippy fixture, root Cargo.lock,
Python cache, temporary logs and 14.9 GiB of build outputs were removed.

Commit staging was attempted with `git add -A`, but shared Git metadata is
read-only: `/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`
could not be created. No commit was created. The complete goal remains active.
