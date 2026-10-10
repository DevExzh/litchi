# Source-backed PPTX related parts

This wave adds the R3 related-part read foundation under
`presentation::related`. The complete unified-operations goal remains active.
Presentation-wide metadata, full chart projection, modern comments, the PPTX
unified adapter, write dispatch, and performance acceptance remain pending.

## Snapshot contract

Source slides acquire layout, notes, legacy comments, and chart snapshots.
Layouts expose their authored name and typed layout kind; notes expose the
primary body placeholder text. Layout and notes snapshots carry independently
retained shape scenes. Comments expose semantic text, timestamp and location;
chart snapshots expose their basic kind, title and legend state. Chart
selection uses `Position` and stable lexical relationship-identifier order;
indexed reads load only the selected chart. Missing optional relationships
return `None`. Chart result lists are bounded before reserving storage.

Snapshots are parsed and validated during acquisition. Their accessors read
immutable retained values after the source owner is dropped. They do not retain
the presentation owner. Retained source payloads intentionally keep their
managed memory reservations until their last snapshot handle is released.
Native part names, relationship identifiers and original payload bytes are
available through the explicit `related::raw` module. Snapshot diagnostics
omit authored content and native package identities.

## Validation and retention

Related reads fence source and execution state around acquisition. Physical
part sizes and XML budgets are checked before MCE preprocessing. MCE output
is capped at the smaller physical/semantic part limit before leaf decoding.
Notes follow
the selected slide backlink, notes master and theme closure, applying the
existing notes grammar without scanning unrelated slides. The selected master
must match the presentation's canonical notes-master declaration. Mixed
strict/transitional required edges and comments outbound edges are refused.
Duplicate singleton
relationships, external selected relationships and mismatched content types
are refused. Layout and notes scenes share the source owner's bounded scene
memo; managed contexts retain no memo entries. Publication has a final fence
and removes the exact published scene when that fence fails.

## Verification

- Full all-feature PPTX suite: 1,398 tests passed, including the first 20
  related-part regressions. Three existing tests were ignored.
- Final focused related-part suite: 22 tests passed after canonical notes
  validation and mixed-conformance regression additions. Coverage includes
  eager/source parity, detached snapshots, raw identity, source changes,
  managed cancellation/budget release, cache reuse/trim, exact publication
  rollback, physical XML limits, bounded MCE expansion, selective chart
  reads, empty notes, strict unknown tokens and invalid relationships.
- Workspace library compilation passed.
- Both PPTX glossary profiles passed with 35 bans each. Glossary and crate
  boundary checker tests passed (1,151); the boundary inventory check passed
  for 66 workspace packages and 267 internal dependency declarations.
- All-feature/all-target PPTX Clippy passed with warnings denied.
- Warning-denied PPTX rustdoc and the new detached-notes doctest passed.

This is a basic chart summary and legacy-comment projection, not full chart
series/axes or modern-comment coverage. Broad format source/semantic options
and incremental semantic execution-budget accounting remain pending. No
paired performance acceptance is claimed.

## Commit and cleanup

Changed Rust sources were formatted; `git diff --check` passed. The changed
and untracked artifact audit passed, and both mandatory staged-file checks
printed nothing.

The three temporary PPTX corpus fixtures copied from the base checkout were
removed with only their newly created empty directories. Generated untracked
lockfiles and Python cache were removed. `cargo clean` removed 24,756 files
(29.5 GiB). Turn scratch was removed; no temporary worktrees were created.
Participating agents reported no additional temporary artifacts.

The rewritten local base remains an ancestor of HEAD, so no additional local
rebase was needed. No remote fetch is claimed. `git add -A` again failed:
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` is on a
read-only filesystem. No commit was created. Records 0051–0061 remain
uncommitted; refresh the stale partial index before an eventual commit.
