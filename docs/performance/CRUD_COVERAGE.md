# Performance CRUD coverage

Date: 2026-08-11

This is a coverage map, not a completion claim. It compares the 117 selectable
benchmark cases with `docs/CRUD_Scenario_Checklist.md`. Generic ZIP/OPC/CFB
substrate measurements do not certify format-semantic CRUD behavior.

| Required scenario | Current status | Measured coverage |
|---|---|---|
| Open and identify format | Partial | ZIP/OPC/CFB plus owned DOC/XLS/PPT/RTF/XLSX and source-backed XLSX open; RTF now covers plain, raw CP-1252, LZFu and a real-producer watermark input; the public PPT root slide-order snapshot has its own measured validation path; no smart-detection handoff case |
| List semantic children without payloads | Partial | XLS/XLSX/ODS sheets, DOC/RTF/DOCX/ODT paragraphs and PPT/PPTX/ODP slides; DOCX section listing remains missing |
| Query one property or named object | Partial | XLS/XLSX/ODS cells, one DOC/RTF/DOCX/ODT paragraph, one PPT shape and one PPTX/ODP slide; broader properties/images remain missing |
| Read one cell/paragraph/slide/image/Part | Partial | XLS/XLSX/ODS cells, DOC/RTF/DOCX/ODT paragraphs, PPT/PPTX/ODP text objects and generic OPC Part; semantic image selection remains missing |
| Scan all cells/paragraphs/slides | Covered for generated native/OOXML/RTF/ODF text corpora | XLS/XLSX/ODS cell scans, including the isolated ODS public-cell sweep; DOC/RTF/DOCX/ODT paragraph enumeration and PPT/PPTX/ODP slide/text enumeration |
| Full text extraction | Covered for generated DOC/PPT/RTF/DOCX/PPTX/ODT/ODS/ODP | Complete deterministic text or row-major cell text is checked; RTF additionally verifies raw CP-1252/LZFu text and the body plus public header-shape projection of a real-producer watermark; ODT consuming block ownership is measured in change 0023; broader real-producer/media-heavy corpora remain missing |
| Semantic conversion to sequential sink | Missing | Package serialization exists; semantic export/conversion does not |
| Create a small document | Partial | Fresh DOC/XLS/PPT plus DOCX/PPTX/ODT/ODS/ODP public authoring; large/streaming creation remains missing |
| Create or append a very large stream | Partial | Large fresh legacy writers accumulate before final output; logical append remains separate and missing |
| Exact no-op edit and commit | Covered for generated DOC/XLS/PPT/RTF/XLSX/DOCX/PPTX/ODT/ODS/ODP | Public semantic transaction plus save/reopen; RTF also proves exact raw CP-1252, LZFu and producer-watermark publication; signed/extension corpora remain missing |
| One semantic edit and save | Covered for generated DOC/XLS/PPT/RTF/XLSX/DOCX/PPTX/ODT/ODS/ODP | Cell/paragraph/shape edit or supported ODP slide append, then save/reopen; ODT, ODS and ODP additionally verify eight exact 2 MiB resources and manifest media types after one paragraph/cell/text-box edit; a separate RTF correctness gate proves real-producer root-shape edit plus checked LibreOffice resave/readback without presenting it as a paragraph benchmark |
| About 1% semantic update and save | Covered for XLSX/DOCX/PPTX generated corpora | Deterministic evenly spaced cell, paragraph and shape changes; DOCX uses one canonical atomic paragraph batch and reopens the package |
| Bulk update matching objects | Missing | No semantic end-to-end case |
| Clear/remove/hide/detach/GC distinctions | Missing | No complete matrix |
| Sanitization and irreversible redaction | Missing | No complete matrix |
| Copy object with dependency closure | Missing | No measured format case |
| Merge and split | Missing | No measured format case |
| Patch encode/apply/invert/merge | Partial | DOCX coalesced replacement correctness covers deterministic durable encode/decode/apply/inverse, but no durable lifecycle timing; broader formats/merge remain missing |
| Validate without mutation | Partial | Opens validate; no distinct validate-only matrix |
| Explicit repair plan | Missing | No general public non-mutating repair-plan API |
| Preserve unknown extension during understood edit | Partial | Targeted OPC raw-copy framing/unknown-member tests, exact untouched opaque ODS-row preservation, and exact raw ODT/ODS/ODP auxiliary/media members during neighboring paragraph/cell/text-box edits; broader format-semantic extension corpora remain missing |
| Replace one low-level Part, preserve the rest | Covered for owned same-topology OPC | Changes 0008/0021/0022 test raw framing, fallback, shared changed-payload ownership, validated local-span movement and matched save behavior; source-backed editing remains missing |

The source/output matrix is also incomplete. Owned bytes and instrumented
`ReadAt` exist for OPC/XLSX, and the deterministic range simulator covers
latency/range effects. RTF and DOCX final serialization accept and test a
forward-only non-seek sink, but this is not semantic conversion or
memory-bounded authoring. Borrowed-byte comparisons, filesystem positional
cold reads, atomic-save timing, PPTX facade streaming output, and non-seek
semantic conversion remain.

## Highest-return next cases

1. Attribute remaining final owner/public-reader work after the accepted native
   XLS editor reuse, DOC batched stream publication, PPT root-open reuse and
   direct PPT text-edit resolver reuse;
   preserve exact-source patch/inverse and the rejected DOC move as independent
   guardrails.
2. Extend native RTF beyond the new raw CP-1252, LZFu and watermark read/no-op
   matrix plus the narrow real-producer shape-text native chain: formatting and
   media semantics, malformed/security cases, more producers and broad edits
   remain open.
3. Separate logical authoring/append time from final serialization and reopen
   for DOCX, PPTX and XLSX.
4. XLSX bulk update plus distinct clear/remove/hide behavior. Direct
   writer-local action regrouping was measured and rejected in change 0030;
   broader coverage must not present that immaterial prototype as a solution.
5. Unknown OOXML extension and media preservation during a known semantic edit.
6. Durable PPTX patch produce/encode/decode/apply/inverse/join/three-way flows,
   including stale-base and conflict cases.
7. PPTX dependency-closure transfer and slide split/removal with charts, media,
   themes and collision names.
8. Validate/security matrix for valid, malformed-within-limits, encrypted,
   macro-enabled, protected and external-link fixtures.
9. Smart detection versus prepared-source reuse. OOXML smart results retain an
   adoptable parsed OPC package; ODF detection/handoff remains unmeasured.
10. Broaden ODF beyond generated text/grid/deck and accepted compact ODT/ODS/ODP
   unchanged-member publication: source-backed selectors, resource-adding and
   structural publication, 1% and bulk edits, unknown extensions, real
   producers, richer media, security and source-backed I/O.
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

Change 0021 removes the measured 4.19 MiB changed-Part handoff copy from owned
same-topology OPC publication. It does not complete source-backed editing,
unknown-extension semantic corpora, or the broader OOXML CRUD rows above.

Change 0022 removes the separate 4.20 MiB post-validation generated-local-span
copy. The complete generated archive is still validated before output, and the
required compression buffer remains; this is not broader OOXML CRUD coverage.

Change 0023 removes two intermediate strings per block only from ODT full-text
extraction. Structured queries retain their former ownership contract, and
source-backed ODF reads, repeated ODT/ODP scans, other-family/structural
publication, bulk edits and broader producer/security coverage remain open.

Change 0024 removes one duplicate CFB index open from public PPT root
slide-order capture. It retains independent live-document, slide-order,
review-history and public-reader validation; it does not imply completion of
PPT edit/publication, real-producer, security, or broader OLE2 CRUD coverage.

Change 0025 adds a distinct XLSX commit-plus-first-read attribution case and
reuses commit-time validation only for bounded changed worksheets with exact
part and style/shared-string identity. The dense-wide handoff was rejected on
peak memory and remains a cold-cache fallback; broader XLSX bulk, structural,
source-backed and preservation scenarios remain open.

Change 0026 adds a distinct direct PPT text-edit attribution case and removes
one repeated editor open from target resolution. The fresh commit editor,
exact source comparison, patch/inverse and complete readback remain; this does
not complete PPT anchors, broad edits, real-producer or security coverage.

Change 0027 adds an aggregation-free public ODS cell-sweep attribution case and
a bounded lazy facade locator for repeated queries. It does not add structural,
bulk, source-backed, real-producer, media, security or publication coverage;
those remain under item 10.

Change 0029 broadens the RTF input matrix without adding or renaming timed case
types: deterministic raw CP-1252, deterministic LZFu and a content-addressed
producer watermark join the plain corpus under explicit capability filters.
It also gates the checked `relsize` source/Litchi/LibreOffice semantic chain.
This is coverage evidence, not an optimization result; formatted/media-heavy,
malformed/security and broad real-producer edits remain open.

Change 0030 attributes the existing XLSX 1% update and save cases to the
writer's nested row/cell action regrouping. An owned forward-stream prototype
improved formal p50 by at most 1.61%, reduced process allocation calls only
0.0623%, and left peak heap flat; it and its prototype-only tests were fully
reverted. Distinct bulk clear/remove/hide, structural edits and any larger
semantic-planning/emission coalescing remain open.

Change 0031 adds a deterministic media-rich ODS one-cell edit/save case and
accepts compact `content.xml` publication that raw-copies exact unchanged ZIP
members. It proves exact opaque/media payloads, manifest media types, no-op and
logical fallback behavior, but does not complete other-family or structural
publication, 1%/bulk edits, real-producer, signature/encryption, source-backed
or broad media semantics.

Change 0034 adds the matching media-rich ODP source-backed text-box edit/save
case and accepts the already shared checked-splice/raw-copy publication path
for content-only rich-object operations. Resource additions and unsupported or
security-sensitive layouts keep the complete rebuild. The case proves exact
media/metadata framing, complete slide and rich-content readback, deterministic
patch/inverse/stale behavior, and material end-to-end latency improvement; it
does not complete structural/bulk edits, repeated queries, real-producer media,
or source-backed positional I/O.

Change 0035 adds the corresponding media-rich ODT paragraph edit/save case and
accepts the common checked-splice/raw-copy path only for content-only paragraph
replacement. It proves raw unchanged-member identity, full semantic/media
readback and patch/inverse/stale behavior. Regenerated `content.xml` above the
common 16 MiB optimization limit explicitly returns to the established ODT
rebuild. Structural/bulk edits, repeated queries, real-producer media and
source-backed positional I/O remain open.

Change 0036 separates the common OLE2 one-stream edit into editor-open,
candidate-publication, final-render and end-to-end attribution cases over four
unchanged 4 MiB streams. It does not add semantic CRUD coverage. An inline
recapture allocation-reuse prototype was reverted because the complete public
operation improved only 2.61% p50/2.30% mean; the cases remain to gate a
materially different final-publication design.

Each new case should use deterministic object positions and digests, separate
semantic work from publication, reopen outputs, verify untouched content, and
record source/sink, allocation and peak-memory behavior in addition to time.
