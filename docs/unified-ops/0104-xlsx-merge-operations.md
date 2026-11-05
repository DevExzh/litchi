# 0104 — XLSX merge geometry and range readback

The unified `merge` property delegates geometry changes to native worksheet
edits. Range targets accept boolean merge/unmerge, a matching A1 range, and
contained-merge `sweep`. Cell and sheet targets accept explicit range lists;
cell unmerge removes only an anchored merge. Nonexact overlapping range
unmerge reports `merge_not_exact` rather than deleting different geometry.

The shared scalar registry admits the merge input grammar and null clearing.
Range readback now reports boolean `merge` from native indexed coverage;
cell readback retains the existing range-text representation, including blank
anchors that have no physical cell record. An XLSX-local
read specification removes the fulfilled range property gap without changing
XLSB. Capability write flags come from the same admission registry.

Merge-only requests stage rectangles without expanding their area into cells.
Selected nodes, parsed lists and staged intents are bounded and deduplicated.
Raw text is byte-bounded before trimming or parsing, so repeated ranges cannot
evade ingress limits. Multi-sheet rename remains ambiguous even when combined
with merge and is refused before staging.
Native grid handles provide merge membership without copied catalog models.
Companion cell or sheet assignments are admitted before one native commit;
combined target/staging charges apply. Merge actions close a pending cell
coalescing run. Source fences and native patches preserve atomic publication,
replay, inverse and exact no-ops.

The native eager and source-backed merge editors now refuse effective merge
additions that intersect typed table ranges, using `TableOverlap`. Existing
exact merges remain exact no-ops. This closes an OfficeCLI parity/safety gap
in the leaf implementation. Follower content, overlapping merges, group
formulas and compatibility-owned or unmodeled payloads remain native guards;
the unified writer does not discard them.

Other XLSX structural families, broader coalescing, format coverage and
Python/CLI acceptance remain open. No performance acceptance claim is made
for this checkpoint.

Validation:

- 911 combined DOCX/XLSX/encryption facade tests/doctests passed.
- 419 minimal-feature tests/doctests passed.
- 37 native merge unit tests and all 21 source-backed merge integration tests
  passed, including eager/source table-overlap refusal and exact no-ops.
- All 15 unified merge regressions passed again after the final textual boolean
  alias adjustment. Coverage includes replay/inverse, blank anchors, huge
  geometry without cell expansion, mixed writes, limits, no-op selection
  accounting, duplicate input bytes and multi-sheet rename refusal.
- Capability/range-view snapshots and 14 generated-delta tests passed.

Cargo ran offline. The rewritten base remains an ancestor, so no rebase replay
is needed against the current base.

Final warning-denied Clippy passed for ops and XLSX, all targets, with
DOCX/XLSX/encryption enabled. Scoped formatting, whitespace checks and the
457-path artifact audit passed; both required staged-file checks printed
nothing.

Independent review confirmed the input-byte and multi-sheet-rename fixes and
found no further blocker in admission, no-op accounting, native guards or
source fencing. Broader merge-selector scaling remains part of the open
performance acceptance work.

The borrowed Clippy fixture, root Cargo.lock, Python cache, temporary logs and
17.4 GiB of build outputs were removed. No temporary worktree was created.

Commit staging was attempted with `git add -A`, but shared Git metadata remains
read-only: `/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`
could not be created. No commit was created. The full goal remains active.
