# 0041 — lazy DOCX canonical publication and atomic save

## Change and policy

The lazy DOCX Package exposed only to_stream for whole-package publication.
It now exposes write_to, save and save_with_durability. The replaced name has
no alias and is banned in both DOCX glossary profiles. Owning Package and
specialized transaction publishers retain their existing names for later waves.

write_to keeps the existing physical no-op copy and retained effective-topology
publication paths. save defaults to Full durability. All durability levels stage
in a sibling temporary file and atomically replace the destination; only file
and directory synchronization differ. Existing permissions and destination
validation continue through the shared OPC atomic helper.

The new guarded OPC helper reserves the fixed 64 KiB staging buffer before
filesystem work, holds the reservation through completion, and checks source
state after flushing and selected file synchronization, immediately before
replacement. Cancellation is checked after that fence as well. Pre-rename
failures preserve the destination and remove the temporary. A failed directory
sync after replacement remains typed Committed. The final source observation
is a fence, not an atomic compare-and-swap against external source writers.

Staging uses the destination owner's execution context, matching publication
workspace/output accounting. An unmanaged destination has unmanaged staging,
even with managed donors. Donor contexts still govern donor reads and retention;
all donor sources and contexts participate in the final source fence. This
accounts for the fixed buffer, not every filesystem allocation or process RSS.

## Verification

Five new atomic tests cover flush/sync/fence ordering, typed SourceChanged at
the final fence, exact and one-under staging admission, cancellation immediately
before rename, reservation release and post-rename Committed errors. Four new
DOCX tests cover exact physical bytes at all three durability levels, effective
edit save over the original path with unchanged compressed opaque payloads,
stale-source refusal and output-budget refusal with destination preservation.

- Full-feature OPC/DOCX/operations unit and integration suites: 3,533 passed,
  two existing ignored (1,103 OPC, 2,041 DOCX, 389 operations).
- Minimal-feature OPC atomic suite: 21 passed.
- Warning-denied OPC/DOCX doctests: 95 passed, 31 existing ignored.
- Workspace all-feature library compilation, including facade/Python: passed.
- Warning-denied OPC/DOCX/operations library Clippy: passed.
- Glossary gate: 11 profiles, 142 banned names, zero complete families; passed.
  All 21 checker unit tests pass.
- Crate boundaries: 66 packages, 267 internal dependencies and 11 existing
  explicit debt items; passed.
- Changed Rust formatting and diff whitespace checks: passed.
- Rewritten base remains 32ad64e320 and is an ancestor of this branch.
- Independent read-only review identified a remaining old-name caller and
  requested final-fence coverage and an explicit donor/destination staging
  policy. The final change addresses all three.

## Remaining work and cleanup

Budgeted to_bytes, owning/specialized output migration, owner consolidation,
semantic/crypto policy migration, other formats, unified consumers and full
acceptance remain active. This is not a complete glossary family or milestone,
and no performance improvement or sustained fuzz campaign is claimed.

The six temporary base-checkout fixture copies listed in record 0035, generated
root lockfile, build output and session logs are removed before commit. No new
worktree was created; base, iWork and pre-existing worktrees are preserved.
