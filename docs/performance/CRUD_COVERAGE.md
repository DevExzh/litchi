# Performance CRUD coverage

Date: 2026-08-10

This is a coverage map, not a completion claim. It compares the 81 selectable
benchmark cases with `docs/CRUD_Scenario_Checklist.md`. Generic ZIP/OPC/CFB
substrate measurements do not certify format-semantic CRUD behavior.

| Required scenario | Current status | Measured coverage |
|---|---|---|
| Open and identify format | Partial | ZIP/OPC/CFB and owned/source-backed XLSX open; no smart-detection handoff case |
| List semantic children without payloads | Partial | XLSX/ODS sheets, DOCX/ODT paragraphs and PPTX/ODP slides; DOCX section listing remains missing |
| Query one property or named object | Partial | XLSX/ODS cells, one DOCX/ODT paragraph and one PPTX/ODP slide; broader properties/images remain missing |
| Read one cell/paragraph/slide/image/Part | Partial | XLSX/ODS cells, DOCX/ODT paragraphs, PPTX/ODP slides and generic OPC Part; semantic image selection remains missing |
| Scan all cells/paragraphs/slides | Covered for generated OOXML/ODF text corpora | XLSX/ODS cell scans, DOCX/ODT paragraph enumeration and PPTX/ODP slide/text enumeration |
| Full text extraction | Covered for generated DOCX/PPTX/ODT/ODS/ODP | Complete deterministic text or row-major cell text is checked; real-producer/media-heavy corpora remain missing |
| Semantic conversion to sequential sink | Missing | Package serialization exists; semantic export/conversion does not |
| Create a small document | Partial | Fresh DOC/XLS/PPT plus DOCX/PPTX/ODT/ODS/ODP public authoring; large/streaming creation remains missing |
| Create or append a very large stream | Partial | Large fresh legacy writers accumulate before final output; logical append remains separate and missing |
| Exact no-op edit and commit | Covered for generated XLSX/DOCX/PPTX/ODT/ODS/ODP | Public semantic transaction plus save/reopen; signed/extension corpora remain missing |
| One semantic edit and save | Covered for generated XLSX/DOCX/PPTX/ODT/ODS/ODP | Cell/paragraph/shape edit or supported ODP slide append, then save/reopen |
| About 1% semantic update and save | Covered for XLSX/DOCX/PPTX generated corpora | Deterministic evenly spaced cell, paragraph and shape changes |
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
latency/range effects. DOCX final package serialization now accepts and tests a
forward-only non-seek sink, but this is not semantic conversion or
memory-bounded authoring. Borrowed-byte comparisons, filesystem positional
cold reads, atomic-save timing, PPTX facade streaming output, and non-seek
semantic conversion remain.

## Highest-return next cases

1. Coalesce DOCX same-structure paragraph replacements so a 1% transaction
   does not rebuild and reparse the complete XML once per paragraph.
2. Separate logical authoring/append time from final serialization and reopen
   for DOCX, PPTX and XLSX.
3. XLSX bulk update plus distinct clear/remove/hide behavior.
4. Unknown OOXML extension and media preservation during a known semantic edit.
5. Durable PPTX patch produce/encode/decode/apply/inverse/join/three-way flows,
   including stale-base and conflict cases.
6. PPTX dependency-closure transfer and slide split/removal with charts, media,
   themes and collision names.
7. Validate/security matrix for valid, malformed-within-limits, encrypted,
   macro-enabled, protected and external-link fixtures.
8. Smart detection versus prepared-source reuse. OOXML smart results retain an
   adoptable parsed OPC package; ODF detection/handoff remains unmeasured.
9. Broaden ODF beyond generated text/grid/deck cases: 1% and bulk edits,
   unknown extensions, real producers, media, security and source-backed I/O.
   iWork is deliberately deferred while the `iwa-*` crates change separately.

Each new case should use deterministic object positions and digests, separate
semantic work from publication, reopen outputs, verify untouched content, and
record source/sink, allocation and peak-memory behavior in addition to time.
