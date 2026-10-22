# 0362rbd — DOCX raw story replay by content

This slice restores standalone replay for the raw story-text codec models.
`story_text::Patch::apply` and `GlossaryBatchPatch::apply` authenticate the
selected member with its selector, case-insensitive part name, byte length,
SHA-256, and exact bytes. They accept an independently opened equal-content
recipient and reject changed content with the crate's canonical
`PatchConflict`. Selector or part mismatches remain `TopologyConflict`.
The SHA-256 value is the immutable fingerprint stored by snapshot admission;
replay compares that authenticated value plus length and exact bytes rather
than hashing the same payload a second time.

The public raw replay path re-admits a changed target under the recipient's
limits and execution context. It reconstructs the selected-story or glossary
layout, carries the recipient source version and lineage, and clears the
author's transaction envelope before returning the raw snapshot. A no-op
returns the admitted recipient snapshot unchanged. `raw::Tree::apply` and
`raw::GlossaryTree::apply` expose this capability without publishing a
package, so a caller can apply a forward patch, its inverse, and a second
edit on a reopened raw model.

The crate-private owner publication helper retains its existing fast path: it
performs the same content proof, then uses the already admitted target bytes.
Only standalone raw replay performs target layout re-admission. The
owner package still performs its complete-artifact, source-freshness,
security, and output publication checks. Owner publication now captures and
checks the current story under the recipient package's retained story limits,
including the final output check; a raw patch authored under broader limits
cannot widen a narrower recipient package. The main-story raw edit also
checks the admitted source length before cloning its transaction XML, so a
tight output policy refuses before allocation.

The old public `ForeignSource` and `ArtifactConflict` variants were removed.
Stale selected content is `PatchConflict`; topology and source-freshness
failures remain distinct. This preserves the error taxonomy while allowing
legal exact-content cross-copy replay, including signed no-op behavior where
the enclosing owner family admits it.

Focused coverage adds independent matching-copy forward/inverse chains for
story text and glossary batches, a stricter recipient-limit story route with
an explicit output-ceiling refusal, broad-producer/narrow-recipient raw and
owner replay routes, changed-content refusal, and cancellation before raw
content comparison for changed and no-op patches. Existing owner tests
continue to cover complete artifact publication and exact inverse restoration.

The focused command was run before the final owner-fast-path factoring; its
terminal output is retained outside the repository. The factoring only moves
the existing selector/content proof into a private helper and leaves public
raw re-admission unchanged. A fresh capped focused run is required before
this slice is treated as fully validated. The original pre-migration
`Snapshot::edit` lifecycle comparison is prepared separately at
`/home/zhuhe/code/litchi-worktrees/.coord/w8-docx-story-text` and is likewise
pending the coordinator's build release.
