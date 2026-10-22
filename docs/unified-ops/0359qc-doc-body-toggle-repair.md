# 0359qc — DOC body-toggle final handoff repair

- **Status:** matched allocator and instruction receipts recorded; native
  qualification remains pending because the host was overloaded.
- **Date:** 2026-10-22.
- **Input:** the ordinary-route attribution in [0359qb](0359qb-doc-body-toggle-qualification.md).
- **Scope:** the existing `litchi-doc` body transaction. No public editing
  family, alternate API, or probe route was added.

## Repair

The deferred CHPX path now has a consuming publication seam for the final
render. `RevisionEditor::finish_unchecked` already owns the editor at that
point, so its private `commit_consuming` helper takes the owned
`WordDocument`, table, and (when present) `Data` vectors before constructing
the lower object editor's `Arc<[u8]>` replacements. This removes the extra
`Vec::clone` allocation and copy that the ordinary one-toggle route otherwise
performs immediately before the `Arc` conversion. The `Vec` to `Arc<[u8]>`
conversion still allocates and copies once, as required by the object-editor
representation.

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

The matched heap, Callgrind, source-assembly, and output hashes for this
candidate are recorded in
[`results/0359qc/qualification.json`](results/0359qc/qualification.json).
The candidate was built with toolchain `1.95.0`, one Cargo job, incremental
compilation disabled, release debuginfo disabled, and the 20 GB systemd cap.
Its selected ordinary route produced the same output SHA as both 0359qb legs.

The candidate's route-local Ir is 1.011632 of the original implementation and
1.003013 of the 0359qb after leg. `Edit::commit` is 1.114649 of the original
and 1.011591 of the 0359qb after leg. The peak total is 0.926057 of the
original and 1.000006 of the 0359qb after leg; the peak tree remains under
`routes::doc_leaf_write` and contains no batch route. These are deterministic
Callgrind and Massif observations, not a claim that the consuming handoff
improved CPU cost. The largest changed callees remain open for the next repair.

No new native ABBA block was run: the host load was
`119.63 138.73 155.23` with unrelated Cargo and rustc work active. The
native blocks and their rejection reasons from 0359qb remain retained. One
initial Callgrind launch also failed before execution because Valgrind could
not create temporary files in shared `/tmp`; the successful capture used the
private DOC `TMPDIR` and is the only profile included in the JSON receipt.
