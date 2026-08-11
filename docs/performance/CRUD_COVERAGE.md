# Performance CRUD coverage

Date: 2026-08-11

This is a coverage map, not a completion claim. It compares the 123 selectable
benchmark cases with `docs/CRUD_Scenario_Checklist.md`. Generic ZIP/OPC/CFB
substrate measurements do not certify format-semantic CRUD behavior.

| Required scenario | Current status | Measured coverage |
|---|---|---|
| Open and identify format | Partial | ZIP/OPC/CFB plus owned DOC/XLS/PPT/RTF/XLSX and source-backed XLSX open; RTF now covers plain, raw CP-1252, LZFu and a real-producer watermark input; the public PPT root slide-order snapshot has its own measured validation path; no smart-detection handoff case |
| List semantic children without payloads | Partial | XLS/XLSX/ODS sheets, DOC/RTF/DOCX/ODT paragraphs and PPT/PPTX/ODP slides; DOCX section listing remains missing |
| Query one property or named object | Partial | XLS/XLSX/ODS cells, one DOC/RTF/DOCX paragraph, an indexed fully validated ODT paragraph, one PPT shape and one PPTX/ODP slide; broader properties/images remain missing |
| Read one cell/paragraph/slide/image/Part | Partial | XLS/XLSX/ODS cells, DOC/RTF/DOCX paragraphs, indexed ODT paragraphs, PPT/PPTX/ODP text objects and generic OPC Part; semantic image selection remains missing |
| Scan all cells/paragraphs/slides | Covered for generated native/OOXML/RTF/ODF text corpora | XLS/XLSX/ODS cell scans, including the isolated ODS public-cell sweep; DOC/RTF/DOCX/ODT paragraph enumeration and PPT/PPTX/ODP slide/text enumeration |
| Full text extraction | Covered for generated DOC/PPT/RTF/DOCX/PPTX/ODT/ODS/ODP | Complete deterministic text or row-major cell text is checked; RTF additionally verifies raw CP-1252/LZFu text and the body plus public header-shape projection of a real-producer watermark; ODT consuming block ownership is measured in change 0023; broader real-producer/media-heavy corpora remain missing |
| Semantic conversion to sequential sink | Missing | Package serialization exists; semantic export/conversion does not |
| Create a small document | Partial | Fresh DOC/XLS/PPT plus DOCX/PPTX/ODT/ODS/ODP public authoring; large/streaming creation remains missing |
| Create or append a very large stream | Partial | Large fresh legacy writers accumulate before final output; logical append remains separate and missing |
| Exact no-op edit and commit | Covered for generated DOC/XLS/PPT/RTF/XLSX/DOCX/PPTX/ODT/ODS/ODP | Public semantic transaction plus save/reopen; RTF also proves exact raw CP-1252, LZFu and producer-watermark publication; signed/extension corpora remain missing |
| One semantic edit and save | Covered for generated DOC/XLS/PPT/RTF/XLSX/DOCX/PPTX/ODT/ODS/ODP | Cell/paragraph/shape edit or supported ODP slide append, then save/reopen; ODT, ODS and ODP verify eight exact 2 MiB resources and manifest media types, while source-backed DOCX/PPTX and narrow XLSX calculation-metadata publication verify eight exact 2 MiB media Parts after one semantic edit; a separate RTF correctness gate proves real-producer root-shape edit plus checked LibreOffice resave/readback without presenting it as a paragraph benchmark |
| About 1% semantic update and save | Covered for XLSX/DOCX/PPTX/ODT generated corpora | Deterministic evenly spaced cell, paragraph and shape changes; DOCX uses one canonical atomic paragraph batch, while ODT coalesces ordinary scalar durable replacements internally; both reopen the package |
| Bulk update matching objects | Missing | No semantic end-to-end case |
| Clear/remove/hide/detach/GC distinctions | Missing | No complete matrix |
| Sanitization and irreversible redaction | Missing | No complete matrix |
| Copy object with dependency closure | Missing | No measured format case |
| Merge and split | Missing | No measured format case |
| Patch encode/apply/invert/merge | Partial | DOCX and ODT coalesced replacement correctness covers deterministic durable encode/decode/apply/inverse, but no durable lifecycle timing; broader formats/merge remain missing |
| Validate without mutation | Partial | Opens validate; no distinct validate-only matrix |
| Explicit repair plan | Missing | No general public non-mutating repair-plan API |
| Preserve unknown extension during understood edit | Partial | Targeted OPC raw-copy framing/unknown-member tests, exact untouched opaque ODS-row preservation, and exact raw ODT/ODS/ODP auxiliary/media members during neighboring paragraph/cell/text-box edits; broader format-semantic extension corpora remain missing |
| Replace one low-level Part, preserve the rest | Covered for owned OPC, narrow source-backed OPC, guarded DOCX main-document semantics, guarded PPTX selected-slide semantics, and XLSX calculation metadata | Changes 0008/0021/0022 test owned raw framing, fallback and payload ownership; change 0037 adds the consuming source-backed one-Part publisher; changes 0039, 0044 and 0046 integrate exact-source DOCX, PPTX and XLSX calculation-metadata transactions while refusing unsafe MCE, signature, transfer/multi-operation, stale/foreign and topology cases before output; general XLSX cell/formula editing remains outside the one-Part closure |

The source/output matrix is also incomplete. Owned bytes and instrumented
`ReadAt` exist for OPC/XLSX, and the deterministic range simulator covers
latency/range effects. The narrow source-backed OPC publisher accepts a
forward-only sink and records complete positional input plus bounded output,
but this is not semantic conversion or memory-bounded authoring. RTF and DOCX
final serialization also accept and test a forward-only non-seek sink.
Borrowed-byte comparisons, filesystem positional cold reads, atomic-save
timing, broad/multi-Part PPTX streaming output, and non-seek semantic
conversion remain.

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
5. Broaden the accepted narrow XLSX calculation-metadata source transaction
   beyond change 0046 only where a complete one-Part semantic closure can be
   proved; general cell/formula/chains still need a wider publication design.
   Broaden DOCX/PPTX beyond changes 0039/0044 with real producers, MCE-aware
   editing, dependency transfers and explicit topology/signature fallback or
   refusal matrices.
6. Unknown OOXML extension and media preservation during a known semantic edit.
7. Durable PPTX patch produce/encode/decode/apply/inverse/join/three-way flows,
   including stale-base and conflict cases.
8. PPTX dependency-closure transfer and slide split/removal with charts, media,
   themes and collision names.
9. Validate/security matrix for valid, malformed-within-limits, encrypted,
   macro-enabled, protected and external-link fixtures.
10. Smart detection versus prepared-source reuse. OOXML smart results retain an
   adoptable parsed OPC package; ODF detection/handoff remains unmeasured.
11. Broaden ODF beyond generated text/grid/deck and accepted compact ODT/ODS/ODP
   unchanged-member publication: source-backed selectors, resource-adding and
   structural publication, 1% and bulk edits, unknown extensions, real
   producers, richer media, security and source-backed I/O. The generated ODT
   1% paragraph case is covered by change 0045; broader bulk and structural
   operations remain.
   iWork is deliberately deferred while the `iwa-*` crates change separately.

The former first item is complete for existing-document ODT transaction
snapshots: change 0014 removes one archive copy and two allocations per
snapshot while retaining the exact no-op, limits, envelope, patch, and readback
contracts. Direct byte ingress and changed publication remain covered by item
11 rather than being implied complete.

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
those remain under item 11.

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

Change 0037 covers the low-level source-backed replacement of one existing OPC
Part without URI, content-type, relationship or topology changes. The matched
case materializes only the selected Part, raw-copies all other physical
members, preserves exact output identity, and refuses signed real changes or
unsupported layouts before output. It does not add DOCX/PPTX/XLSX semantic
transactions, topology changes, signature policy, real-producer/media matrices
or atomic filesystem publication.

Change 0039 integrates that publisher with exact-source DOCX main-document
transactions over a fixed 16 MiB media-rich package. It materializes only the
main Part, raw-copies every other physical member, fully reopens the result and
retains byte-exact no-op/signature/source-version behavior. MCE-normalized
documents and paragraph transfers are deliberately refused; PPTX/XLSX,
topology changes, real producers, encrypted input and atomic filesystem
publication remain open.

Change 0044 integrates that publisher with an exact-source PPTX selected-slide
transaction over a fixed 200-slide, eight-media package. It materializes the
mandatory presentation root and selected slide, raw-copies the other 227
logical Parts, fully reopens the result, preserves every untouched Part and
media payload, and retains byte-exact no-op/signature/source-version behavior.
MCE-normalized slides, more than one edit operation, topology changes,
real-producer matrices, encrypted input and atomic filesystem publication
remain deliberately outside the contract.

Change 0046 integrates the publisher with an exact-source XLSX
calculation-metadata transaction over a fixed 12-Part, eight-media package. It
materializes only `xl/workbook.xml`, raw-copies the other 11 Parts, fully
reopens the result, verifies the typed calculation state, and preserves every
untouched Part and media payload. MCE projection, changed signed sources,
stale/foreign workbook closures, topology changes and partial sinks retain
typed refusals. Cells, formulas, cached results, styles, shared strings,
relationships and calculation-chain ownership remain outside the capability.

Change 0047 makes the existing ODT one-paragraph benchmark a direct public
indexed query. It retains only the requested structured paragraph while
scanning and validating the complete XML, including malformed trailing input
and all resource limits. It does not provide positional source I/O, early XML
termination, an ODP selector, or broader real-producer/media query coverage.

Change 0049 makes the existing ODP one-slide benchmark a direct public indexed
query. It retains completed semantic content only for the requested slide while
resolving all transition styles and validating content through EOF. It proves
public/full-list parity, out-of-range behavior, late semantic failure and style
inheritance failure. It does not provide positional ZIP/XML I/O, early
termination, repeated-query caching, or broader real-producer/media selection.

Change 0040 removes repeated UTF-8 scalar decoding from ordinary RTF text
delimiter discovery without changing the existing CRUD surface. It measures
plain, raw CP-1252 and LZFu opens at medium and large, retains the 25-row
transport/producer smoke, and preserves exact no-op, opaque syntax, limits,
patch/inverse and complete readback contracts. Formatting/media-heavy,
malformed/security and broader real-producer edits remain open.

Change 0043 measures and fully reverts direct RTF decoded-body ownership. It
adds no CRUD surface and keeps the existing parser/model handoff. The retained
test covers a lossy incomplete Shift-JIS body byte plus byte-exact immutable
publication; the broad prototype's plain/LZFu regression and the owned-only
variants' sub-threshold instability are negative evidence, not support claims.

Change 0048 keeps the same generated RTF paragraph-edit CRUD surface but
removes a repeated source clone/lexer from eligible changed commits. The
retained range is proven during the initial full parser preflight; empty,
ambiguous, binary, non-ASCII and LZFu inputs retain the established fallback or
refusal, and every changed candidate still receives complete parse/readback.
The capability-bounded 63-record transport/producer smoke remains green. This
does not add formatted/media editing, conversion, security repair or broader
real-producer mutation coverage.

Change 0041 removes three archive-sized copies from the existing ODT
changed-operation compactness audit without adding a CRUD surface. The
media-rich paragraph edit/save still parses both packages, audits compact XML,
reopens the final document, reads back the paragraph and media, and verifies
patch/inverse and stale-source behavior. Ordinary open/edit/no-op guards remain;
the unchanged sub-microsecond exact no-op segment's +39 ns p50 movement is
disclosed. Structural/resource-adding ODF edits, broader real producers and
repeated ODT/ODP reads remain open.

Change 0042 removes the remaining archive-sized envelope-classification copy
from changed ODT commits without adding a CRUD surface. The media-rich
paragraph edit/save still validates the ZIP, parses the manifest, classifies
encryption/signature state, audits compact XML, reopens the final document,
reads back paragraph/media content, and verifies patch/inverse and stale-source
behavior. Open/no-op/edit guards remain; the untouched large exact no-op
segment's +152 ns p50 movement is disclosed. Broader structural/resource edits,
real producers and repeated ODT/ODP reads remain open.

Each new case should use deterministic object positions and digests, separate
semantic work from publication, reopen outputs, verify untouched content, and
record source/sink, allocation and peak-memory behavior in addition to time.
