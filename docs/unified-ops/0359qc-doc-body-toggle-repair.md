# 0359qc — DOC body-toggle final handoff repair

- **Status:** source repair committed; build and performance qualification
  remain pending a quiet build slot.
- **Date:** 2026-10-22.
- **Input:** the ordinary-route attribution in [0359qb](0359qb-doc-body-toggle-qualification.md).
- **Scope:** the existing `litchi-doc` body transaction. No public editing
  family, alternate API, or probe route was added.

## Repair

The deferred CHPX path now has a consuming publication seam for the final
render. `RevisionEditor::finish_unchecked` already owns the editor at that
point, so its private `commit_consuming` helper moves the owned `WordDocument`,
table, and (when present) `Data` buffers into the lower object editor's
`Arc<[u8]>` replacements. This removes the final stream-buffer copies that
the ordinary one-toggle route otherwise performs immediately before its
validated render.

The reusable `RevisionEditor::commit` path is unchanged: it flushes pending
CHPX, authorizes the edit, clones replacements, and publishes them through the
failure-atomic `put_streams_shared` operation. The consuming helper is called
only by a consuming finish, so a failed final render cannot mutate the staged
editor that the body transaction still owns. Protection authorization,
retained-render limits, signed-source checks, and exact no-op handling remain
on the same finish and body-commit paths.

Because `finish_unchecked` is used on a clone for candidate validation and on
the consumed editor for final publication, the transfer is safe for
read-after-write and interleaved operations: serialized-stream dependencies
still flush pending CHPX before they reopen the editor. Dropping an edit still
discards its candidate. The deferred batch path therefore keeps its one-time
FKP rebuild and package publication.

## Validation boundary

`cargo fmt --all -- --check` and `git diff --check` passed for the repair.
The focused mixed-format/table test reached the linker, but the linker
reported `Cannot create temporary file in /tmp/: No space left on device`
while unrelated host builds were active. No Rust diagnostic or test failure
was produced. The test and ordinary-route qualification must be rerun with
the repository's one-job capped build after the host build window is quiet.

The matched heap, Callgrind, native-attempt, source-assembly, and output
hashes already recorded in 0359qb remain the evidence for the pre-repair
implementation. This record makes no performance claim for 0359qc until a
new after executable has been built from this source commit and compared with
the same protected before executable and harness.
