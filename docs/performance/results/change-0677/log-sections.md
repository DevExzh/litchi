# Log sections for change 0677

## For `HOTSPOTS.md`

## 0677 — marked XML no longer fails publication at a truncated declaration

[0677](0677-xml-publication-bom-offsets.md) closes 0651 row 16's publication
auditor offset defect. Both slice and streaming auditors now address the
physical bytes of UTF-8-marked XML. All 16 marked XML members in three real
packages pass the source audit. The streaming path also stops reconstructing
the slice auditor's formerly incorrect spans. No speed claim is made.

## For `GOAL_AUDIT.md`

## 0677 — a valid encoding marker is preserved through publication

[0677](0677-xml-publication-bom-offsets.md) turns 0650's byte-zero declaration
refusal into successful marked-part publication and byte-identical readback.
The DTD regression still refuses before any output; malformed and repeated
markers remain rejected. This is a correctness fix under 0652's standing
trade-off, separate from the managed DOCX transaction's offset correction.

## For `REPORT.md`

## 0677 — XML audit transport parity includes valid UTF-8 markers

[0677](0677-xml-publication-bom-offsets.md) records a failing-before,
passing-after declaration witness, all 16 marked XML members across three real
packages, tiny-chunk slice/stream parity, and a public OPC replacement
publication test. Input-byte limits and reports include the marker; XML token
limits exclude it. The removed streaming scratch is not promoted to an
allocation or latency claim. `performance_claim: none`.

## For `ADR_COMPLIANCE.md`

## 0677 — physical offsets corrected without changing validation boundaries

[0677](0677-xml-publication-bom-offsets.md) preserves ADR 0006's typed
structural, encoding, DOCTYPE and budget checks, and the OPC audit-before-output
boundary. Exactly the leading marker is encoding framing; a second marker
remains character data. No signature changes, unsafe code, dependency or
ambient I/O is introduced. The streaming memory envelope remains conservative.
