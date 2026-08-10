# Performance CRUD coverage

Date: 2026-08-10

This is a coverage map, not a completion claim. It compares the 44 selectable
benchmark cases with `docs/CRUD_Scenario_Checklist.md`. Generic ZIP/OPC/CFB
substrate measurements do not certify format-semantic CRUD behavior.

| Required scenario | Current status | Measured coverage |
|---|---|---|
| Open and identify format | Partial | ZIP/OPC/CFB and owned/source-backed XLSX open; no smart-detection handoff case |
| List semantic children without payloads | Partial | XLSX sheets, including zero-request source-backed listing; no DOCX section or PPTX slide listing |
| Query one property or named object | Partial | XLSX first cell; lower-level one-Part/stream reads are not semantic property queries |
| Read one cell/paragraph/slide/image/Part | Partial | XLSX cell and generic OPC Part; paragraph/slide/image cases missing |
| Scan all cells/paragraphs/slides | Partial | XLSX full and narrow scans; DOCX/PPTX missing |
| Full text extraction | Missing | No end-to-end format case |
| Semantic conversion to sequential sink | Missing | Package serialization exists; semantic export/conversion does not |
| Create a small document | Partial | Fresh DOC/XLS/PPT writers; no OOXML creator case |
| Create or append a very large stream | Partial | Large fresh legacy writers accumulate before final output; logical append remains separate and missing |
| Exact no-op edit and commit | Partial | XLSX patch/commit/save; generic exact OPC save is not a semantic edit |
| One semantic edit and save | Partial | XLSX one-cell only |
| About 1% semantic update and save | Partial | XLSX only |
| Bulk update matching objects | Missing | No semantic end-to-end case |
| Clear/remove/hide/detach/GC distinctions | Missing | No complete matrix |
| Sanitization and irreversible redaction | Missing | No complete matrix |
| Copy object with dependency closure | Missing | No measured format case |
| Merge and split | Missing | No measured format case |
| Patch encode/apply/invert/merge | Partial | XLSX in-memory commit only; no durable lifecycle timing |
| Validate without mutation | Partial | Opens validate; no distinct validate-only matrix |
| Explicit repair plan | Missing | No general public non-mutating repair-plan API |
| Preserve unknown extension during understood edit | Partial | Targeted OPC raw-copy framing/unknown-member tests; no format-semantic extension corpus |
| Replace one low-level Part, preserve the rest | Covered for owned same-topology OPC | Change 0008 tests and matched save benchmark; source-backed editing remains missing |

The source/output matrix is also incomplete. Owned bytes and instrumented
`ReadAt` exist for OPC/XLSX, and the deterministic range simulator covers
latency/range effects. Borrowed-byte comparisons, filesystem positional cold
reads, atomic-save timing, and non-seek semantic conversion output remain.

## Highest-return next cases

1. DOCX paragraph scan, full text, one-paragraph and 1% edit/save through
   `litchi_docx::Package` and document edit APIs.
2. PPTX slide scan/full text plus one-shape and 1% edit/save through the opened
   presentation transaction APIs.
3. Separate logical authoring/append time from final serialization and reopen
   for DOCX, PPTX and XLSX.
4. XLSX bulk update plus distinct clear/remove/hide behavior.
5. Unknown OOXML extension and media preservation during a known semantic edit.
6. Durable PPTX patch produce/encode/decode/apply/inverse/join/three-way flows,
   including stale-base and conflict cases.
7. PPTX dependency-closure transfer and slide split/removal with charts, media,
   themes and collision names.
8. Validate/security matrix for valid, malformed-within-limits, encrypted,
   macro-enabled, protected and external-link fixtures.
9. Smart detection versus prepared-source reuse; iWork has an opaque
   `PreparedSource`, while generic OOXML currently has no reusable handoff.

Each new case should use deterministic object positions and digests, separate
semantic work from publication, reopen outputs, verify untouched content, and
record source/sink, allocation and peak-memory behavior in addition to time.
