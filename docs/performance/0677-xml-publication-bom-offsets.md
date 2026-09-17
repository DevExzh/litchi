# 0677: XML publication audits address BOM-marked input correctly

Status: implemented correctness follow-up. `performance_claim: none`.
OLE2 and OOXML remain the active priority; ODF optimization is deferred and
iWork is excluded.

## Problem and authority

Queue row 16 of [0651](0651-queue-refresh-after-the-second-wave.md) requests the
publication-audit offset fix witnessed by
[0650](0650-docx-editor-byte-order-mark-admission.md), finding 3. The slice
auditor used `Reader::buffer_position()` to slice the original input. quick-xml
excludes the initial three-byte UTF-8 marker from those positions, so the
first declaration was audited as the marker plus a truncated declaration and
refused as `Malformed` at byte zero.

The streaming auditor requires a correction too: contrary to an implication of
0650's description, its marker bookkeeping intentionally reconstructed the
slice auditor's incorrect spans. The old `reader_matches_slice_bom_rejection`
test explicitly required that refusal. Fixing only one transport would break
their parity.

This follows the correctness-first rule of
[0652](0652-owner-decisions-for-the-third-wave.md). It removes no structural
check or compactness rule and makes no new owner decision.

## Change

The slice auditor adds the marker length to parser positions when addressing
raw input and reporting diagnostics. It continues to pass the original input
to quick-xml, which consumes exactly the first marker: a second marker remains
character data and is refused outside the document element.

The streaming guard exposes the complete initial marker to quick-xml, including
when a token budget is less than three bytes. Consuming that prefix charges
input bytes, but not XML-token bytes. The guard captures only the lexical token
for validation. Reported offsets include the marker; the old marker-shifted
capture buffer and history arrays are removed. The published conservative
memory envelope remains unchanged and continues to bound the smaller scratch
requirement.

When a streaming input-byte budget is below three, the prefix cannot yet be
classified as a complete marker; existing streaming precedence can report a
small token limit before the total-byte limit. This rule is unchanged.

Neither path rewrites the input. Aggregate byte limits and `Report::bytes()`
include all three marker bytes. Token, event, depth, attribute and text limits
still describe XML content. Unmarked input takes the same validation path and
keeps its offsets.

## Compatibility

Valid marked XML is now accepted by `verify`, `verify_authored`,
`verify_source`, and both streaming audit entry points. This also admits an
OPC replacement part that carries a marker. Marker-bearing failures now name
the physical input offset, rather than a position three bytes early. No public
signature changes. Malformed XML, duplicate markers outside the root, invalid
UTF-8, DTDs, multiple roots and budget violations remain typed refusals. The
OPC publication plan still audits before writing any output.

The managed DOCX transaction has its own offset defect; its correction belongs
to [0670](0670-docx-parser-residues.md), not this auditor change.

## Evidence and validation

The new admission regression fails on the prior code with the exact old
`Malformed { offset: 0, detail: "invalid XML declaration boundary" }` witness.
After the correction, the XML auditor suite passes, including tiny-chunk
slice/stream parity, repeated and truncated markers, token limits, invalid
encoding offsets, deterministic malformed mutations, and unchanged authored
whitespace policy. Additional tests assert marker-inclusive byte accounting,
physical token-limit offsets, DTD and malformed-document refusals.

A public OPC regression replaces a part with marked XML, publishes and reopens
it byte-identically, then replaces it with a marked DOCTYPE and verifies that
publication refuses with an empty sink. A corpus test audits all 16 marked XML
members in three real packages, including 0650's `alt-chunk-header.docx`.
The member census and commands are retained in the
[packet](results/change-0677/README.md).

This is admission and diagnostic evidence, not a performance measurement.
Removing the streaming marker scratch is visible in the implementation; no
latency, allocation-count or RSS saving is claimed.
