Four log paragraphs for change 0650, for the coordinator to merge into the
shared program logs. Each block below is written in the style of the newest
section of the file it names.

## For `HOTSPOTS.md`

## 0650 — the DOCX editor's admission hole was three bytes wide, and it is not the only site

Record: [0650](0650-docx-editor-byte-order-mark-admission.md).

**Change [0638](0638-facade-and-ordinary-save-selectors.md)'s admission fact is
characterized, and it is a parser defect, not a contract.** 0638 found
`Package::document_mut()` refusing `alt-chunk-header.docx` with *"invalid DOCX
XML: syntax error: tag not closed"* although `Package::open` admits it and
`Package::save` republishes it, over 600 retained samples, and could change no
file under `crates/`. The extent is now counted: **63** DOCX-family fixtures
live under `test-data/` — not hundreds; `.docm` and `.dotm` are absent —
`Package::open` admits **55** and refuses **8**, all eight at the OPC or shared
OOXML layer and none in the DOCX editor; of the 55, `document_mut()` admits
**53** and refuses **2**; the reader refuses **none**, so **there is no fixture
the editor admits and the reader refuses**. The mechanism is quick-xml 0.41.0
removing a leading UTF-8 byte order mark from the slice (`reader/slice_reader.rs:257`)
**without counting it in `Reader::buffer_position`**, so every caller that mixes
reported offsets with slices of its own buffer reads **three bytes low for the
whole document**: `DocumentBody::from_xml`'s preserved body ranges begin and end
three bytes early and each child is truncated inside its own closing tag. The
reader is immune only because `DocumentPart::from_part` parses the
markup-compatibility-normalized output, which carries no mark — the probe
asserts `reader_bom part=true view=false` on every marked input. The minimal
witness is **five elements**: a marked document with one paragraph, refused
before and accepted after; its unmarked twin is accepted on both legs. **Exactly
one of the 63 fixtures is marked**, which is why 0638 saw one file; the corpus is
LibreOffice and POI output, and Word marks its parts. The fix is one split in
`DocumentBody::from_xml` plus a carry through `MutableDocument`'s serializer,
and value identity is total: `diff` of two 630-line censuses is **three changed
lines, all refusal text on the one marked fixture**, with all 55 no-op, 53
editor-touched, 54 managed-route and 50 edited published SHA-256 digests
identical. **Three further sites of the same defect class are witnessed and left
open**: `crates/xml-minifier/src/audit.rs:1319` (the authored-XML publication
audit, which therefore refuses a *regenerated* marked part with *"malformed XML
at byte 0: invalid XML declaration boundary"* while its own streaming auditor
handles the mark explicitly), `crate::document::Snapshot` in the managed
transaction (refusing a marked dense body with `expected '</w:bo<w:p>'`
identically on both legs), and — a different defect — the writer's section parse
at `writer/section/codec/xml.rs:409`, which rejects ISO Strict's `w:pgSz
w:w="612pt"` because it admits only `ST_UnsignedDecimalNumber`. Editing a marked
DOCX end to end still does not work; this removes the first of four obstacles
and makes the diagnosis accurate. `performance_claim: none`; no timing taken.
OLE2/OOXML remain active; ODF is deferred until completion and iWork excluded.
[Record and limitations](0650-docx-editor-byte-order-mark-admission.md);
[retained evidence](results/change-0650/README.md).

## For `GOAL_AUDIT.md`

## 0650 — a refusal that was wrong rather than conservative, and the three it was hiding

Record: [0650](0650-docx-editor-byte-order-mark-admission.md).

`docs/GOAL.md` puts correctness and lossless preservation above speed and says a
typed refusal must never be traded for a partial result. The refusal this change
removes was not a conservative one: it told the caller that well-formed markup
was malformed, on bytes the same crate's reader parses successfully. The audit
distinction worth recording is that **removing it admits nothing that was
previously refused for a reason** — the fixture that prompted the investigation,
`alt-chunk-header.docx`, is **still refused**, now with the accurate
`body-final section properties are not the final body child`, and the managed
document transaction was already giving exactly that message on the same file on
the unmodified base, so the two routes now agree rather than one of them
inventing a syntax error. Preservation is served in the same move: the mark is a
byte the producer wrote, so it is carried into the preserved prefix and back out
through the serializer rather than discarded, and the `touch_save` census aspect
— open, acquire `document_mut()`, save, change nothing — publishes **exactly the
digest the plain no-op publishes** on the newly admitted inputs. Four audit rows
this change opens rather than closes. First, **the end-to-end route is still
broken and this record does not claim otherwise**: a *regenerated* marked part is
refused by the authored-XML publication audit in `xml-minifier`, which has the
same offset defect in its slice path while its streaming path handles the mark
explicitly — out of scope by the brief, reported with a witness. Second, the
managed document transaction has the same defect and is **not** fixed here,
because that route carries source-backed splice reservations and exact-source
byte proofs that need their own value-identity argument. Third, the editor and
the reader genuinely **disagree about `w:sectPr` placement**: ECMA-376's
`CT_Body` makes `w:altChunk` after the body-final `w:sectPr` schema-invalid, so
the editor's refusal is defensible, but the reader admits the file and reads its
text, and which side should move is frozen with two witnesses that differ only in
the order of the last two body children. Fourth, `strict.docx` shows a second,
unrelated editor-versus-reader split — the writer's section parse admits only
`ST_UnsignedDecimalNumber` and so rejects ISO Strict's `612pt`, while the reader
opens the file, reads its five paragraphs and its one section, and the managed
route publishes it; that is a feature gap in the mutable section model, identical
on both legs, and belongs in its own record. **Corpus reach is the honest
weakness**: one marked fixture in 63, so the newly admitted class is represented
by constructed packages rather than by real producer output. No latency,
instruction, allocation or RSS improvement is claimed and no timing was taken.
OLE2/OOXML remain active; ODF is deferred until completion and iWork excluded.

## For `REPORT.md`

## 0650 — a three-byte mark that made a valid document look broken

Record: [0650](0650-docx-editor-byte-order-mark-admission.md).

`litchi-docx`: Microsoft Word puts a three-byte invisible marker at the very
start of the XML files inside a `.docx`, saying "this text is UTF-8". Litchi's
DOCX *reader* has always coped with it. Litchi's DOCX *editor* did not: the XML
parser it uses quietly skips that marker but does not tell the caller it did, so
the editor's idea of where every paragraph, table and section began was three
bytes behind the truth. Every piece of the document it tried to keep verbatim was
snipped three characters short — cut off inside its own closing tag — and the
editor then reported the document as malformed. Change 0638 had noticed this on
one file and could not chase it; this change chases it. The extent, measured over
**all 63** Word-family test documents in the repository: the package opener
accepts 55 and the editor accepts 53 of those, refusing two; the reader refuses
none. **Exactly one** of the 63 carries the marker, which is why only one file
ever showed the fault — the repository's documents come from LibreOffice and
Apache POI, which do not write it, while Word does. The smallest document that
reproduces it is five tags long. The fix is a single line that sets the marker
aside before parsing and puts it back on the way out, so nothing is silently
dropped. Nothing else moved: over 630 recorded observations on those 63 files,
**three lines changed and all three are the wording of a refusal on the one
marked file** — every saved document is byte-for-byte identical, and the file
that started all this is **still refused**, now with an accurate explanation (its
section settings are not the last thing in the body, which the format does not
allow) instead of a false claim that its XML is broken. What this does **not**
do: editing a marked Word document and saving it still fails, at a later step, in
a different crate that has the same three-byte blind spot — that is reported with
a witness for its own change, not fixed here, along with two other places the
same pattern appears. Three new tests; `cargo fmt`, clippy, 1,468 `litchi-docx`
tests, rustdoc, the 266-test feature-bearing facade suite and the two tool
packages that link this crate all pass. `performance_claim: none`; no timing was
taken and none is claimed. [Change and
limitations](0650-docx-editor-byte-order-mark-admission.md); [retained
evidence](results/change-0650/README.md).

## For `ADR_COMPLIANCE.md`

## 0650 — compliant; ADR 0006's "preserve by default" decides what to do with three bytes

Record: [0650](0650-docx-editor-byte-order-mark-admission.md).

Change 0650 is compliant and adds no ADR question of its own, but it is decided
by one. **ADR 0006 (preserve by default; readers preserve real-world quirks)**:
the leading UTF-8 byte order mark is a byte a real producer wrote, so the fix
splits it off for the duration of the parse and returns it at the head of the
preserved prefix, and `MutableDocument::to_xml`/`to_xml_with_rels` carry it
across the compactor the parser would otherwise have eaten. Discarding it would
have been one character shorter and would have made the end-to-end save work; it
is not what ADR 0006 asks for, and the record says plainly that the route
therefore still stops at a later typed refusal. **ADR 0003 (edits are
source-checked, reversible; a refusal is typed, never a partial result)** is
preserved in both directions: no error variant is added, removed or relocated;
the two refusals that change are `Error::Xml`/`Error::InvalidFormat` before and
after; the save stays atomic, so where a caller now meets a publication refusal
instead of a parse refusal, no partial package is written. **Error identity is
otherwise exact** — over 63 fixtures and ten documented aspects each, the only
lines that differ between legs are three refusal *texts* on the single marked
fixture, and that fixture is still refused. **Bounded resources are honoured**:
`MAX_XML_BYTES` in `alt::codec::validate_xml` and `MAX_SCAN_DEPTH` /
`MAX_SCAN_NODES` in `namespace::scan_word_element_ranges` are untouched and
neither is relocated, widened or restated; they now see a buffer three bytes
shorter on marked input. No new `unsafe` (the crate remains
`#![forbid(unsafe_code)]`), no weakened malformed-input defence, no ambient I/O,
no hidden Rayon pool, no executor or lock exposed, no new dependency, no new or
changed public type or signature, and no allocation added on the unmarked path.
**Validation still does not mutate**: `validate_section_placement` and
`final_section_properties` run on every parse as before, on the same elements in
the same order, and now receive the ranges they were always meant to receive.
Two compliance rows stay open and are stated plainly. First, the **authored-XML
publication audit** in `crates/xml-minifier` refuses a regenerated marked part
(`audit.rs:1319` slices with `Reader::buffer_position` offsets; `check_declaration`
at `:1542` then rejects the boundary) while the same crate's streaming auditor
handles the mark explicitly — a contract inconsistency, since a *pristine*
passthrough member carrying the same bytes is republished without complaint. It
is out of this change's scope by the brief and is reported, not touched. Second,
the **editor and the reader disagree about body-final `w:sectPr` placement**;
ECMA-376's `CT_Body` supports the editor, the reader is the lenient one, and the
resolution — which would change where an appended paragraph lands — is frozen
here with witnesses rather than decided. `performance_claim: none`. OLE2/OOXML
remain active; ODF is deferred until completion and iWork excluded.
