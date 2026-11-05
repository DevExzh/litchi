# 0362ul87ac — OOXML properties replay and staging repair

Status: source-only follow-up to 0362ul87ab.  No Cargo, rustdoc, native
qualification, performance measurement or final gate was run for this repair.

## Semantic staging

The shared metadata transaction now distinguishes semantic assignment from
explicit graph removal:

* `set_custom_props` compares the decoded collection.  Setting an equal empty
  collection is an exact no-op, preserving the present raw XML and signed
  package state; a changed collection still creates, replaces, or removes the
  family as appropriate.
* `remove_custom_props` is the separate removal intent.  It removes a present
  custom-properties part and relationship even when tolerant decoding left an
  empty projection or recorded skipped properties.  An absent graph remains a
  no-op, and signed removal is refused before publication.

The intent is available on the common `Plan`/`Edit` and DOCX, XLSX, PPTX and
both eager/source-backed XLSB metadata edits.  DOCX's package bridge passes the
intent through the shared patch planner.  Existing DOCX raw-empty behavior is
covered both ways: semantic setting keeps the raw part byte-for-byte, while
the explicit method removes its graph.

An explicit metadata set now records core presence when the source has no
core part, even when the requested core projection is `Metadata::default()`.
The resulting creation patch and its inverse agree with the wrapper's
`None`/`Some(default)` endpoints.

## Replay and fidelity

Eager and source-backed XLSB metadata no-op replay now authenticates the
recorded property-family scope.  Equal owners accept the copy, foreign
recorded bases return `PatchConflict`, and a no-op captured from a revised
source-backed owner is checked against that owner's current composed revision.

DOCX metadata regressions assert that core-only edits preserve `Application`,
`Template`, and opaque application XML bytes.  Combined core/application/
custom edits retain those fields and opaque bytes, publish all requested
families, and still replay/invert exactly.

## Source-only checks

`git diff --check` and scoped `rustfmt --edition 2024 --check` pass for all
changed Rust sources.  The coordinator should run the affected common, DOCX,
XLSX, PPTX and XLSB tests, including signed/no-op/removal, dirty custom graph,
core-presence and source replay cases.
