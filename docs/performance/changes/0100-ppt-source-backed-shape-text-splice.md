# Change 0100: source-backed same-length PPT shape-text splice

Date: 2026-08-14

Status: Accepted as correctness-only API coverage. No performance result is
claimed.

## Scope

`litchi_ppt::text_edit` now has an additive immutable source-backed owner for
one existing shape text atom selected by `(slide_position, shape_position)`.
It accepts an exact no-op or a replacement with the same encoded byte length
as the existing `TextBytesAtom` or `TextCharsAtom`, then publishes the change
through the protected common CFB same-length splice boundary. Length-changing
text continues to use the ordinary owned editor.

The transaction binds the source version, length, complete artifact
fingerprint, live slide persist identity, selected atom offset and expected
bytes. Forward and inverse application re-resolve the semantic owner and
reject stale or foreign sources. Publication performs a complete composed-CFB
reopen and selected-shape readback before exposing output; unrelated streams
are preserved, exact no-ops reuse the source identity, and sequential-sink
failures report accepted progress.

## Fail-closed closure

The narrow owner refuses or rejects:

- signed, encrypted, protected, macro-bearing and embedded-storage sources;
- root/`PP97_DUALSTORAGE` stream-pair ambiguity or cross-topology pairing;
- unsupported slide, drawing, shape, textbox, text/style or encoding
  ownership;
- multiple `PPDrawing` owners, duplicate shape/textbox/text atoms, malformed
  OfficeArt roots and trailing partial records;
- active, orphaned or malformed VBA metadata; and
- length changes, overlapping/out-of-bounds splices, stale versions,
  fingerprint changes and finite-limit violations.

The source-backed resolver currently materializes complete bounded copies of
the `PowerPoint Document` and `Current User` streams to reuse the existing
live persist and slide directory implementation. The common publisher also
retains full-artifact integrity checks. This change therefore establishes a
validated equal-length replacement-staging and physical-splice API, not
proportional selected-range source I/O, lower latency, lower allocation count,
bounded total memory, peak-heap/RSS improvement or cold-filesystem behavior.

## Verification

Generated-writer regressions cover semantic reopen, exact no-op, inverse,
unrelated-stream preservation, partial sinks, exact/one-under limits,
protection and macro refusals, invalid text/header metadata, multiple complete
OfficeArt roots, trailing records and cross-topology stream pairing. The full
library suite, warning- and deprecation-denied focused suite, strict library
Clippy, rustdoc warning gate and independent adversarial review are required
for this tranche. Real-producer, media-rich, fragmented CFB and matched release
ABBA/resource evidence remain open.
