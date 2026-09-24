# XLSB Data Model and connections graph tests

This test-only batch verifies the combined Data Model and External Data
Connections ownership boundary through public package/workbook APIs. It adds
positive binding and case-insensitive binding controls, missing/duplicate/orphan
connections ownership, wrong content type, external and outbound relationships,
missing/ambiguous connection names, and rejected graph edits that preserve exact
saved workbook bytes.

MS-XLSB section 2.1.7.24 defines the singleton workbook-owned connections part,
its content type, and its relationship constraints. Section 2.1.7.35 delegates the
Model part to MS-XLDM and MS-XLSX section 2.1.6. The existing model owner resolves
each table's connection name to exactly one workbook connection. These tests
exercise that composed boundary rather than testing the two owners independently.

Fixtures use the public synthetic workbook writer and deliberately inert opaque
model bytes. This is package-graph evidence, not validation of an XLDM payload,
native Excel producer compatibility, refresh, credentials, calculation, or
external access. No production implementation or performance claim changes.
The native XLSB Data Model producer-evidence gap remains open.

`validation.json` records the focused checks, exact source inputs before/after,
toolchain, and log hashes. The new integration target is run together with the
existing Data Model target; passing those tests is scoped evidence only.
