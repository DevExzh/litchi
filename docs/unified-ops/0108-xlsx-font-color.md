# 0108 — XLSX font color writes

The unified cell/range writer now admits `font.color` and its `color` alias
through the shared spreadsheet color grammar. Named, hex, CSS and typed RGB
colors map to native ARGB; spreadsheet theme colors retain their theme slot.
Both retain checked tint. Unsupported bases and transforms refuse atomically,
and null removes the authored color.

Native `FontChange::color` consumes the existing typed `FontColor`, discarding
parsed lexical markup from authoring ownership and equality. Native RGB,
theme, indexed, automatic and default bases use the same bounded stylesheet
editor. Switching bases removes mutually exclusive attributes; unrelated
markup survives. Semantic repeats preserve source bytes.

Color assignments participate in native joins, replay/inverse history and
immutable source fences. Font changes remain barriers to payload coalescing.
Ordinary readback retains OfficeCLI parity: theme tint is exposed, whereas
RGB tint is preserved natively but omitted from the facade color projection.
Broader formatting, format coverage and full performance acceptance remain open.

Validation to date:

- 929 combined DOCX/XLSX/encryption facade tests and doctests passed offline.
- Both plain and encrypted capability snapshots passed their six-test suites.
- The OfficeCLI capability delta was regenerated and checked; all 14 generator
  tests passed.
- Review caught and corrected same-base quote preservation, absent tint
  insertion, and combined font-child insertion ordering. Final native runs
  passed all 32 stylesheet and 32 font-focused tests, including both empty-tag
  forms and repeat no-ops. All 34 cell-write regressions passed again after
  the final native change.

The rewritten base remains an ancestor of this branch; no rebase replay is
needed against the current base.


Warning-denied all-target Clippy passed for ops and XLSX with combined
DOCX/XLSX/encryption features. Scoped formatting and whitespace checks passed.
The changed-path artifact audit found no oversized evidence or compiled
binaries; both mandatory staged-file checks printed nothing.

The borrowed Clippy fixture, root Cargo.lock, Python cache, temporary logs
and 19.1 GiB of build outputs were removed. No temporary worktree was created.

Staging failed because shared Git metadata is read-only: Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`. No commit
was created. Verified changes remain in the working tree; the full goal
remains active.
