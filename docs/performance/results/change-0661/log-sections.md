# Log sections for change 0661

## For `HOTSPOTS.md`

### 0661 — the accepted C2′ lazy OPC seam is implemented and retained

Record: [0661](../../0661-opc-lazy-part-decode.md). Authority: accepted ADR
0030, change 0652 decision 3, and queue row 3 of change 0651. Owned OPC
packages now retain one source archive and defer typed-part payload inflation
until `get_part`, `get_part_mut`, `main_document_part`, `part_by_reltype`, or
`try_iter_parts` requests it. `iter_parts` exposes metadata only. The
eager-versus-lazy differential has 1,342 `MATCH` rows and zero `MISMATCH`
rows; exact no-op decodes zero parts in the 336-fixture read-set. No timing or
registered performance claim is made.

## For `GOAL_AUDIT.md`

### 0661 — correctness and bounded failures remain ahead of the deferred work

Record: [0661](../../0661-opc-lazy-part-decode.md). The package invariant
forces a deferred payload before a `&dyn Part` leaves `litchi-opc`; the
infallible metadata iterator cannot reach payload bytes. Each deferred decode
records its first typed error, and the consumer audit now treats only
`PartNotFound` as optional absence. The corrupt-member witness proves that
first-access refusal is stable and cannot become an empty generated payload.
Exact source no-op copying remains an intentional preservation contract, and
changed members are forced and validated before publication.
The follow-up malformed-deferred witnesses extend this to graph publication:
PPTX master/layout authoring rolls back after late decode failure, and XLSB
threaded graph operations stay bound to the resolved root workbook.

## For `REPORT.md`

### 0661 — deterministic read-set evidence, no performance claim

Record: [0661](../../0661-opc-lazy-part-decode.md), evidence packet
[`results/change-0661`](README.md). Across 336 fixture rows, the admitted
catalog is 5,077 parts and 45,562,463 inflated bytes across 334 successful
opens. The exact-no-op route decoded 0 parts/0 bytes; one-part reblob decoded
334 parts/389,581 bytes; the current XLSX hide route published 33 fixtures and
decoded 216/550 parts and 6,044,263/8,274,037 bytes among those publications.
These are deterministic access counters and differential results, not timing,
RSS, allocation, or claim-registry measurements.
The follow-up tests add correctness evidence only and do not alter these
read-set totals.

## For `ADR_COMPLIANCE.md`

### 0661 — ADR 0030's accessor, migration, and source-retention gates are covered

Record: [0661](../../0661-opc-lazy-part-decode.md). The accepted ADR 0030
contract is implemented: owned-source deferred payloads, fallible forcing
accessors, metadata-only `iter_parts`, fallible `try_iter_parts`, stable
first-access failures, and retained exact-source publication. The focused lazy
tests, the 437-test OPC suite, workspace check, and 1,342-row differential
support the record. The packet explicitly withholds the old 0610 one-part XLSX
claim and all timing/RSS/allocation claims; broad release measurement remains
outside this change.
The audit follow-up additionally verifies staged PPTX graph authoring, forced
theme fallback, root-relationship XLSB resolution, and propagation of root
workbook decode failures.
