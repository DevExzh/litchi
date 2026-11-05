# 0103 — XLSX row and column layout assignments

The unified writer now delegates row height, column width, visibility, outline
level and collapsed-state assignments to native worksheet layout edits. A
direct path can address an implicit dimension; scoped queries preserve the
existing read semantics, including stored column-record enumeration and row
table-column predicate resolution. Each successful call retains one native
commit and its patch, preserving source snapshots and native no-ops.

The static row/column registries own coercion, aliases, constraints and null
policy. Height uses the existing point-based length vocabulary, width uses
native character-unit numbers, and outline levels are checked in 0..=7.
`outline` accepts `outlineLevel` and `group` aliases; readback continues to use
`outlineLevel`. Null resets explicit height/width, shows hidden dimensions,
clears outline levels and expands collapsed dimensions. Capability properties
are derived from these registries, including writable aliases.

Selectors deduplicate sheet/axis/index identities and enforce target limits
before staging. Mixed families refuse before mutation. The staged-byte charge
is 64 bytes per target plus fixed typed payload sizes for supplied assignments,
including resets; this is a logical charge, not measured allocator use.
Assignments are admitted once. Source revision checks bracket native work;
direct cell/range paths bypass the dimension route's additional checks.

Known cell/row/column protection guards now report `document_protected`;
modeled capability fences report `unsupported_capability`, preserving their
native typed causes. They no longer fall through to `corrupt_file`.

Whole-axis cell payloads, table-driven property keys, additional formatting,
row/column insertion/deletion, unit-qualified column width conversion and
broader batch coalescing remain explicit follow-ups. This checkpoint makes no
performance acceptance claim. The full all-format/Python/CLI goal remains open.

Validation:

- 895 combined DOCX/XLSX/encryption tests/doctests passed, including nine
  source/eager dimension regressions and registry/error-category tests.
- 419 minimal-feature tests/doctests passed.
- Capability snapshots and 14 generated OfficeCLI delta tests passed.
- Independent read-only review found no remaining actionable routing, source
  freshness, bounded-target, alias/reset or native-semantics defect.

Cargo ran offline. No leaf XLSX implementation changed. The rewritten base
remains an ancestor, so no rebase replay is needed against the current base.

Warning-denied Clippy passed for ops and XLSX, all targets, with
DOCX/XLSX/encryption enabled. Scoped formatting, whitespace checks and the
451-path repository artifact audit passed. Both required staged-file checks
printed nothing.

The borrowed Clippy fixture, root Cargo.lock, Python cache, temporary logs and
14.9 GiB of build outputs were removed. No temporary worktree was created.

Commit staging was attempted with `git add -A`, but shared Git metadata remains
read-only: `/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`
could not be created. No commit was created. The complete goal remains active.
