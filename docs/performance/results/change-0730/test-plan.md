# 0730 DOC validated-render handoff test and corpus plan

Status: independent qualification plan for the bounded DOC pilot. This file
does not make a performance claim and does not expand the format scope to
iWork. The production handoff is tested through the public
`litchi_doc::body_text` API; direct `RevisionEditor` APIs remain outside this
retention contract.

## Contract under test

`TransactionLimits::with_max_retained_render_bytes` adds a finite ceiling to
the body transaction policy. `Edit::retained_render_bytes()` reports
`Some(capacity)` only when the edit owns the already-produced rendered `Vec`,
and reports `None` otherwise. `Edit::release_retained_render()` drops that
allocation without changing the edit's semantic state. The tests must verify
the reported value is the allocation capacity, so a retained value is charged
by the policy rather than by an inferred serialized length.

The retained render is an operation result owned by the current edit state,
not a general CFB cache. The candidate must therefore satisfy all of these
properties:

* The default policy is finite (the pilot default is 8 MiB), and a caller can
  set a smaller or larger ceiling without changing the source bytes or the
  transaction's other limits.
* A zero, undersized, or otherwise exceeded ceiling silently drops the
  rendered allocation and later recomputes through the existing finish path.
  It must not create a new refusal or change the published bytes.
* A retained value is the `Vec` already produced by the validated batched
  render. The owner does not clone it, serialize it to scratch, or retain a
  parsed replacement in its place.
* `Some(capacity)` is never greater than the configured ceiling. An exact
  ceiling retains the same allocation; an over-ceiling case reports `None`.
  The exact boundary is calibrated from the same fixture, source identity,
  replacement, and build so the test does not assume `Vec::capacity() ==
  Vec::len()`.
* Every effective mutation invalidates a previous handoff and either replaces
  it with a render for the new state or falls back to recomputation. A failed
  mutation leaves the editor's semantic state and its prior retained-state
  report unchanged. A true no-op does not alter the semantic state.
* Explicit release reports `None`, leaves all later semantic and physical
  results unchanged, and makes the next commit use the ordinary recomputation
  path. Dropping or rolling back an edit also releases the owner-side
  allocation.
* Final strict-owner validation, public-reader validation, exact no-op source
  sharing, and patch construction remain mandatory. A retained handoff cannot
  bypass the final reopen/readback gate.

## Fixture and producer matrix

The first correctness gate uses the two exact 0728/0730 DOC cases and the
same length-changing paragraph-zero replacement used by the paired probe:

| Case | Source | Qualification role |
| --- | --- | --- |
| `docnohf` | `test-data/ole/doc/NoHeadFoot.doc` | small Word 97 source; primary retention-boundary and repeated-edit case |
| `docfloat` | `test-data/ole/doc/FloatingPictures.doc` | larger Word source with unrelated drawing/opaque streams; preservation and fallback case |

The integration tests add at least one already-supported newer-generation
source, `test-data/ole/doc/documentProperties.doc`, and one source with
auxiliary content such as `ThreeColHeadFoot.doc` or `commented-table.doc`.
They select a plain, non-empty ordinary body paragraph and leave auxiliary
stories, tables, fields, drawings, directory metadata, CLSIDs, and unknown
streams untouched. Existing tests establish the expected FIB generation for
the multi-generation sources; this handoff test checks that the same source
and generated output fully reopen rather than asserting a physical layout.

A generated `writer::Writer` document supplies the format-producer control.
It contains multiple ordinary paragraphs, including non-ASCII text and a
length-changing target, so the handoff is exercised against a current Litchi
producer in addition to imported fixtures. The generated control is a
correctness corpus member only; it does not establish producer breadth.

Before adoption, extend the matrix with every existing DOC body transaction
fixture whose target is inside the proven plain-text closure, including a
header/table/field source for equal-length auxiliary edits. Keep drawing,
tracked-text, protected, malformed, and password-protected sources as
negative or refusal controls. No fixture from an iWork family belongs here.

## Required test groups

### Retention boundaries and observability

For each primary fixture, create an independent uncapped/reference commit and
record its exact output bytes, semantic paragraph text, stream inventory, raw
directory oracle, and patch. Run the same edit with these ceilings:

1. zero: edit succeeds, reports `None`, and publishes exactly the reference;
2. below the calibrated retained capacity: edit succeeds, reports `None`,
   and publishes exactly the reference;
3. exactly the calibrated capacity: edit reports `Some(capacity)` equal to the
   ceiling and publishes exactly the reference; and
4. above the capacity: edit reports `Some(capacity)` below the ceiling and
   publishes exactly the reference.

Calibration must use the same source bytes, selected target, replacement,
limits, and binary identity. If capacity is not stable enough to exercise an
exact boundary, the implementation needs a deterministic test seam or the
case is a gate failure; the test must not weaken the boundary to a length
comparison. Assert that over-budget retention is a recomputation fallback,
not `Error::Refused`, and that no output or semantic witness changes.

### Invalidation, no-ops, and atomicity

After a successful staged replacement, test an effective second replacement,
formatting change, transfer, and each supported body mutation that can be
staged in the same `Edit`. The old report must not describe the new state;
the final retained bytes, when present, must belong to the latest state.
Call a semantic no-op (the current text or current formatting value) and
verify that it neither creates a false new handoff nor changes the eventual
output. Attempt a refused replacement (operation/replacement limit,
structural target, missing target, or a deliberately unsupported dependency)
after a retained handoff. The refusal must leave the current semantic output,
retained report, and subsequent commit unchanged. Also cover rollback and
explicit release.

Use both one-operation and successive-operation transactions. Compare every
successful result with an independently recomputed baseline, not only with a
candidate produced by the same edit object. This catches stale-token reuse
and accidental dependence on operation order.

### Publication, patches, composition, and transfers

For every changed fixture result:

* reopen the committed bytes through the strict DOC owner and the public
  reader, then compare selected paragraph text and the existing stream,
  metadata, CLSID, raw-directory, and unknown-stream witnesses;
* assert exact no-op source sharing and byte identity for a no-op edit;
* apply the forward in-memory patch to the exact source and compare the
  resulting snapshot with the commit;
* apply the inverse to the committed snapshot and compare exact source bytes;
* apply the forward patch to a different real fixture and assert a conflict
  without mutating that source;
* prepare disjoint edits, compose and commit them, then compare the result to
  a fresh sequential baseline; and
* prepare/apply an inert text transfer and reopen the result. If a transfer
  follows an earlier handoff, verify the earlier retained state is invalidated
  and the transfer's output equals an independently recomputed transfer.

The same checks run after release and after a retention fallback. They must
prove that retention affects only allocation lifetime, never patch meaning,
source guards, inverse behavior, or semantic readback.

### Policy meets across equal-byte values

Two separately opened snapshots with equal bytes are the same DOC lineage even
when their `TransactionLimits` differ. Prepared edits and transfer plans must
carry the policy of the planning value, while donor policy remains irrelevant
to a read-only transfer. The integration cases therefore cover:

* disjoint prepared edits from broad and strict equal-byte snapshots joining
  successfully, with all four output limits equal to their componentwise
  minimum, including a zero retained-render ceiling;
* a transfer planned from an equal-byte strict receiver applying to a broad
  receiver and producing the same strict minimum;
* a rejected overlapping join returning recoverable prepared work without
  changing the accepted composition's policy, and a rejected foreign transfer
  leaving both the held-render report and receiving policy unchanged; and
* forward and inverse patch application to an equal-byte source with stricter
  limits, plus a three-way merge whose branch before/after policies are all
  intersected before publication.

Policy assertions use the existing four limit getters. They do not expose or
depend on private retention-token representation.

### Producer and negative controls

The generated Writer document runs the full retention matrix and
forward/inverse checks. Refusal controls cover a non-existent paragraph,
structural/drawing content, tracked text where available, and a replacement
over the ordinary transaction limit. Malformed or protected sources continue
to fail at their existing validation/refusal boundary. A source mismatch must
remain a typed conflict, not a recomputation opportunity.

## Evidence and adoption gates

The focused test file may establish correctness and ownership invariants only.
It must not report a speedup, allocation reduction, or broad producer claim.
Root-owned validation remains responsible for the workspace gates, existing
native fixture oracle, and paired default-feature before/after measurements.
The pilot is eligible for adoption only if all focused tests pass on both
primary fixtures and the generated producer, every fallback and boundary is
byte/semantic identical, retained capacity never exceeds its policy, and no
source, patch, inverse, composition, or transfer gate regresses. Any failure
keeps the candidate and raw evidence for review and blocks a production
policy change.
