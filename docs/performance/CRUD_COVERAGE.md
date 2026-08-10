# Performance CRUD coverage

Date: 2026-08-11

This is a coverage map, not a completion claim. It compares the 106 selectable
benchmark cases with `docs/CRUD_Scenario_Checklist.md`. Generic ZIP/OPC/CFB
substrate measurements do not certify format-semantic CRUD behavior.

| Required scenario | Current status | Measured coverage |
|---|---|---|
| Open and identify format | Partial | ZIP/OPC/CFB plus owned DOC/XLS/PPT/RTF/XLSX and source-backed XLSX open; no smart-detection handoff case |
| List semantic children without payloads | Partial | XLS/XLSX/ODS sheets, DOC/RTF/DOCX/ODT paragraphs and PPT/PPTX/ODP slides; DOCX section listing remains missing |
| Query one property or named object | Partial | XLS/XLSX/ODS cells, one DOC/RTF/DOCX/ODT paragraph, one PPT shape and one PPTX/ODP slide; broader properties/images remain missing |
| Read one cell/paragraph/slide/image/Part | Partial | XLS/XLSX/ODS cells, DOC/RTF/DOCX/ODT paragraphs, PPT/PPTX/ODP text objects and generic OPC Part; semantic image selection remains missing |
| Scan all cells/paragraphs/slides | Covered for generated native/OOXML/RTF/ODF text corpora | XLS/XLSX/ODS cell scans, DOC/RTF/DOCX/ODT paragraph enumeration and PPT/PPTX/ODP slide/text enumeration |
| Full text extraction | Covered for generated DOC/PPT/RTF/DOCX/PPTX/ODT/ODS/ODP | Complete deterministic text or row-major cell text is checked; real-producer/media-heavy corpora remain missing |
| Semantic conversion to sequential sink | Missing | Package serialization exists; semantic export/conversion does not |
| Create a small document | Partial | Fresh DOC/XLS/PPT plus DOCX/PPTX/ODT/ODS/ODP public authoring; large/streaming creation remains missing |
| Create or append a very large stream | Partial | Large fresh legacy writers accumulate before final output; logical append remains separate and missing |
| Exact no-op edit and commit | Covered for generated DOC/XLS/PPT/RTF/XLSX/DOCX/PPTX/ODT/ODS/ODP | Public semantic transaction plus save/reopen; signed/extension corpora remain missing |
| One semantic edit and save | Covered for generated DOC/XLS/PPT/RTF/XLSX/DOCX/PPTX/ODT/ODS/ODP | Cell/paragraph/shape edit or supported ODP slide append, then save/reopen |
| About 1% semantic update and save | Covered for XLSX/DOCX/PPTX generated corpora | Deterministic evenly spaced cell, paragraph and shape changes; DOCX uses one canonical atomic paragraph batch and reopens the package |
| Bulk update matching objects | Missing | No semantic end-to-end case |
| Clear/remove/hide/detach/GC distinctions | Missing | No complete matrix |
| Sanitization and irreversible redaction | Missing | No complete matrix |
| Copy object with dependency closure | Missing | No measured format case |
| Merge and split | Missing | No measured format case |
| Patch encode/apply/invert/merge | Partial | DOCX coalesced replacement correctness covers deterministic durable encode/decode/apply/inverse, but no durable lifecycle timing; broader formats/merge remain missing |
| Validate without mutation | Partial | Opens validate; no distinct validate-only matrix |
| Explicit repair plan | Missing | No general public non-mutating repair-plan API |
| Preserve unknown extension during understood edit | Partial | Targeted OPC raw-copy framing/unknown-member tests plus exact untouched opaque ODS-row preservation during a neighboring cell edit; broader format-semantic extension corpora remain missing |
| Replace one low-level Part, preserve the rest | Covered for owned same-topology OPC | Change 0008 tests and matched save benchmark; source-backed editing remains missing |

The source/output matrix is also incomplete. Owned bytes and instrumented
`ReadAt` exist for OPC/XLSX, and the deterministic range simulator covers
latency/range effects. RTF and DOCX final serialization accept and test a
forward-only non-seek sink, but this is not semantic conversion or
memory-bounded authoring. Borrowed-byte comparisons, filesystem positional
cold reads, atomic-save timing, PPTX facade streaming output, and non-seek
semantic conversion remain.

## Highest-return next cases

1. Attribute remaining final owner/public-reader work after the accepted native
   XLS editor reuse and DOC batched stream publication; preserve exact-source
   patch/inverse and the rejected DOC move as independent guardrails.
2. Extend native RTF coverage beyond the accepted ordinary parser-state and
   ASCII transport batching changes to formatting/media, compressed and
   legacy-code-page input, malformed/security cases, real producers and broad
   edits.
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
9. Smart detection versus prepared-source reuse. OOXML smart results retain an
   adoptable parsed OPC package; ODF detection/handoff remains unmeasured.
10. Broaden ODF beyond generated text/grid/deck cases: source-backed selectors,
   unchanged ZIP-member publication, structural, 1% and bulk edits,
   unknown extensions, real producers, media, security and source-backed I/O.
   iWork is deliberately deferred while the `iwa-*` crates change separately.

The former first item is complete for existing-document ODT transaction
snapshots: change 0014 removes one archive copy and two allocations per
snapshot while retaining the exact no-op, limits, envelope, patch, and readback
contracts. Direct byte ingress and changed publication remain covered by item
10 rather than being implied complete.

The former native DOC/XLS/PPT baseline item is complete in change 0015. Changes
0016 and 0017 accept the first XLS and DOC publication follow-ups without
removing final validation. Change 0018 accepts same-topology ODS row-local
publication and exact untouched opaque-row preservation. Change 0019 accepts
ordinary RTF parser-state work elimination and records a rejected, reverted
ODS package-adoption candidate. Change 0020 accepts RTF ASCII transport
batching and records a rejected, reverted ODT final-document adoption whose
common read guard regressed. None claims broad native-format or ODF CRUD,
real-producer, security, or preservation coverage.

Each new case should use deterministic object positions and digests, separate
semantic work from publication, reopen outputs, verify untouched content, and
record source/sink, allocation and peak-memory behavior in addition to time.
