# Change 0102: range-aware PPT shape-text selection

Date: 2026-08-14

Status: Accepted as correctness and selector-stage source-I/O coverage. No
end-to-end latency, allocation, total-memory, peak-heap/RSS, or cold-filesystem
result is claimed.

## Scope

The source-backed equal-length PPT shape-text owner introduced in change 0100
now resolves `(slide_position, shape_position)` without materializing complete
`PowerPoint Document` and `Current User` streams. The resolver reads the
bounded Current User prefix, the live backward UserEdit/PersistDirectory chain,
record headers through the live Document container, the presentation
SlideListWithText metadata, and the one selected Slide record. The retained
owned resolver remains a test-only differential oracle.

Source opening still fingerprints the complete CFB artifact. Same-length
splice planning, candidate validation, semantic readback, and sequential
publication likewise retain their complete-source integrity checks. This
change therefore reduces selector-stage stream materialization only; it does
not make the complete transaction or save proportional to the selected atom.

## Bounds and refusal closure

Every UserEdit and PersistDirectory header, length, version, instance, backward
offset, and physical range is checked. Persist runs use fallible storage and a
shared finite expansion ceiling; duplicate or overlapping identifiers within
one directory refuse, while legitimate newest-wins identifiers across
historical generations remain supported. Header-only traversal charges record
and depth limits without pretending that skipped payloads were copied;
materialized metadata and selected records charge the copied-payload budget.

The existing macro/storage, encryption/signature/protection, stream-topology,
slide/shape/text ownership, exact source/version/fingerprint, stale/foreign,
same-encoded-length, candidate-reopen, inverse, no-op, and partial-sink gates
remain unchanged.

## Verification

A generated six-slide/four-shape corpus compares every resolved target field
with the owned oracle. After CFB open, the range resolver performs 33
instrumented reads totaling 1,595 bytes while the logical `PowerPoint
Document` stream is 9,329 bytes; its largest request is smaller than the
stream. This is a deterministic test counter, not a release performance
measurement.

Adversarial regressions cover active and duplicate macro owners, duplicate
presentation SlideList owners, trailing selected-slide records, UserEdit
cycles, forged historical UserEdit/PersistDirectory headers, forward or
overlapping directory topology, persist-map expansion at exact/one-under
limits, and header-only versus materialized copy accounting. Focused and full
PPT suites, warning- and deprecation-denied checks, strict library Clippy,
rustdoc, rustfmt, diff checks, and independent adversarial review gate the
change. Real-producer, fragmented-CFB, high-latency, allocation/RSS, and
matched release evidence remain open.
