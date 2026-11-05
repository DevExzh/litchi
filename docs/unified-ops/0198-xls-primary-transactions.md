# 0198 — XLS primary package transactions

## Change

`Package::edit()` and `edit_with_options()` now stage the existing checked
cell-values engine. `Commit` owns the canonical package/patch/diagnostics triad,
with `into_package()` and the `(Package, Patch)` split. `Package::apply()` checks
an exact source image and retains the recipient's policies; `Patch::inverse()`
restores exact bytes. Dropping an edit leaves its source unchanged.

The primary edit delegates numeric and typed values, supported cell insertion
and removal, row/column edits, sheet rename, styles, rich text and formulas.
Bounded scalar transfers accept a Package donor, acquire a checked snapshot,
and publish staged changes only after the donor's source fence passes. The
native dependency/protection refusals remain authoritative. No generic mutable
projection or alternate publication path is exposed.

Both changed and no-op commits retain the effective cell snapshot, including
all seven object-capture bounds and all three CFB bounds. Later default edits
reuse that profile; explicit options can only tighten it. The first default
positional edit also derives its CFB profile from the retained source. Changed
positional packages acquire an owned positional CFB and preserve their complete
source-backed read options, including global/SST and query limits. Eager packages
rebuild their reader under the already admitted candidate profile and retained
leniency/compatibility. Neither route substitutes eager SourceLimits defaults
for a positional profile. Tightened input limits are checked before copying
an existing image during edit admission or exact patch replay.

Detached edit commits check source identity before and after leaf publication,
including no-ops. Signed and encrypted sources remain refused. Source read
failures and truncation now perform the final identity check before classifying
the failure, preserving typed SourceChanged errors across byte, stream and
atomic-save publication. Destination files remain untouched on failed saves.

The leaf's fixed 4,096-change staging ceiling now reports a typed
StagedChangeLimit. The primary family maps it to an inline operations limit;
the facade classifies both forms as limit_exceeded and retains typed causes.
Physical source and CFB failures retain their existing typed error families.

The earlier owner review also identified option propagation gaps outside the
new edit module. Explicit lazy reader materialization now retains CFB limits
and leniency. Explicit facade eager XLS input limits can exceed the default
up to the native hard cap, and positional CFB opening retains an explicitly
selected physical input profile.

## Verification

The combined XLS and XLS/encryption facade suite passes 2,223 tests and
doctests, with one ignored. The 18 owner integration tests cover exact commit,
read-after-write, atomic save/reopen, inverse and replay, stale patches, no-ops,
dropped edits, signed/encrypted refusal, source changes, truncation, retained
narrow and wide profiles, source query bounds, transfer, and failure-atomic
staging-limit refusal. The final native rerun passes 1,747 tests and doctests
with one ignored, including the pre-copy admission change.

Warning-denied all-target Clippy and documentation pass for XLS and the
XLS/encryption facade. The rustdoc glossary audit checks 18 banned names and
one transaction family. The boundary audit passes for 66 packages and 268
internal dependencies, retaining 11 documented debt items. Independent review
found no remaining transaction-bridge blockers. Both repository lockfiles are
unchanged; the rewritten base remains an ancestor of this branch. The performance
harness library and binaries compile with `xls-source-attribution` enabled,
using an external lockfile and build directory.

No benchmark acceptance is claimed. The candidate keeps a validated
edit snapshot, but changed read projections still require a CFB/semantic reopen;
no zero-reparse, zero-copy or total-memory guarantee is claimed.

## Remaining scope

Builder creation, unified XLS node/read/write/batch/publication dispatch,
canonical full-sheet selectors and read coverage, and unified source/semantic
option envelopes remain open. The 4,096-change ceiling must be reconciled with
the 10k batch target before claiming that gate. The explicit positional
materialize_eager byte bound remains specific to that read fallback; focused
primary edits use their retained object/CFB profiles. Broader Tier 2/Tier 3,
Python/CLI, performance and full-goal acceptance remain open.
