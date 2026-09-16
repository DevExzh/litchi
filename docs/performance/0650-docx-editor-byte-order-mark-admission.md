# 0650: the DOCX editor read every offset three bytes low on a byte-order-marked main document, and called the result malformed

Status: retained, a correctness fix in `litchi-docx`. `performance_claim: none`
— this record carries a deterministic admission census over every DOCX fixture
in the repository and a constructed minimal witness, not a claim-registry entry
and no timing. Three files under `crates/litchi-docx/src/writer/doc/` changed,
plus that module's test file.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

Change [0638](0638-facade-and-ordinary-save-selectors.md) reported an admission
fact it could not chase, because no file under `crates/` was allowed to change
there: `litchi_docx::Package::open` accepts
`test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk-header.docx`
and `Package::save` republishes it, but `Package::document_mut()` — the
documented way to edit an opened DOCX — refuses it with

```text
invalid DOCX XML: syntax error: tag not closed: `>` not found before end of input
```

It was one fixture, stable over 600 retained samples, and uncharacterized. This
record censuses the extent, finds the mechanism with a constructed witness,
fixes the half that is a parser defect, and freezes the four questions the
mechanism uncovered.

## What was changed

Three source files and one test file, all under
`crates/litchi-docx/src/writer/doc/`:

* `doc.rs` — one private constant, `BYTE_ORDER_MARK`, with the reason it exists.
* `doc/package.rs` — `DocumentBody::from_xml` splits a leading UTF-8 byte order
  mark off its input before it parses, and returns it at the head of the
  preserved prefix.
* `doc/model.rs` — `MutableDocument::to_xml` and `to_xml_with_rels` call a new
  private `compact_preserving_byte_order_mark` instead of
  `compact_changed_document_xml` directly, so the mark survives compaction.
* `doc/tests.rs` — three tests and two shared inputs.

The tracked diff is +132 / −5 across the four files.

### The mechanism

`quick-xml` removes a leading UTF-8 byte order mark from its source before it
emits the first event, and **does not count it in `Reader::buffer_position`**.
In version 0.41.0 this is `XmlSource for &[u8]`:

```rust
fn detect_encoding(&mut self) -> io::Result<Option<DetectedEncoding>> {
    if let Some(detected) = crate::encoding::detect_encoding(self) {
        *self = &self[detected.bom_len() as usize..];
        return Ok(Some(detected));
    }
    Ok(None)
}
```

(`reader/slice_reader.rs:257`, reached from the `ParseState::Init` arm of
`read_event_impl!` in `reader/mod.rs:269`.) The slice advances; the position
counter does not. Any caller that mixes reported offsets with slices of *its
own* buffer therefore reads three bytes low for the whole document.

`DocumentBody::from_xml` is such a caller three times over. It scans the buffer
for alternative-format anchors (`alt::scan`), for the active block ranges
(`parts::document_part::active_block_ranges`), and then walks it again with an
`NsReader`, recording `reader.buffer_position()` before and after every event
and slicing `xml[start..end]` to preserve each body child verbatim. On a marked
part every one of those ranges begins three bytes early and **ends three bytes
early**, so each preserved child is truncated inside its own closing tag.
`validate_section_placement` then re-parses such a fragment and quick-xml
reports the truncation it was handed: *tag not closed*.

The reader never sees this, because it does not parse the part's raw bytes.
`DocumentPart::from_part` calls `visible_document_xml`, which runs the
markup-compatibility pass and returns its normalized output; that output carries
no mark. The probe asserts this directly — over the whole corpus and the
constructed variants, every marked input reports `reader_bom part=true
view=false`, so the offsets the reader mixes with its buffer are offsets into a
buffer whose mark is already gone. **The same bytes, a valid document, the
reader right and the editor wrong**: the condition for a fix rather than a
frozen contract.

### The fix

`DocumentBody::from_xml` now begins:

```rust
let (byte_order_mark, xml) = match xml.strip_prefix(super::BYTE_ORDER_MARK) {
    Some(rest) => (super::BYTE_ORDER_MARK, rest),
    None => ("", xml),
};
let bytes = xml.as_bytes();
```

One split, before anything reads the buffer. Both scanners and the event loop
then address the same bytes quick-xml addresses, so no offset arithmetic
changes and no range moves. The mark is prepended to the preserved prefix after
`ensure_writer_namespace_declarations` has run on the markup, and
`MutableDocument::to_xml`/`to_xml_with_rels` carry it across the compactor,
which would otherwise consume it the same way.

The carry is deliberately **not** placed inside `compact_changed_document_xml`.
That function has two other callers — `package::package::transfer` compacts one
rewritten *paragraph* and `document::transaction` compacts an already projected
main document — and neither is on this route. Keeping the carry in
`MutableDocument` leaves both byte-for-byte unchanged, which the census
confirms: the managed-route digest is identical on all 63 fixtures.

## Why it is sound

**No invariant, limit, defence or error type moved.** The split is a three-byte
prefix test on a `&str`; it adds no `unsafe` (the crate is `#![forbid(unsafe_code)]`),
no dependency, no allocation on the unmarked path, and no new public item. The
two limit-bearing scanners, `MAX_XML_BYTES` in `alt::codec::validate_xml` and
`MAX_SCAN_DEPTH`/`MAX_SCAN_NODES` in `namespace::scan_word_element_ranges`, now
see a buffer three bytes shorter on marked input and are otherwise untouched;
neither limit is relocated, weakened or restated.

**Exact no-ops stay exact.** `Package::save` without an edit never builds a
`MutableDocument`, so it cannot be affected; the census proves it, with all 63
`noop_save` digests identical between legs. The stronger statement is the
`touch_save` aspect — open, call `document_mut()`, save, change nothing — which
on the newly admitted marked variants publishes **exactly the digest the plain
no-op publishes** (`9326225933e1…0302` and `c8b441736b40…26bde`). Acquiring the
editor does not perturb the bytes.

**The refusal is still typed and still total.** Where `from_xml` now succeeds
and a later stage refuses, the refusal is a typed `Error`, the save is atomic,
and no partial package is published. Nothing was traded for a partial result.

**ADR reading.** ADR 0003 requires edits to be source-checked and refusals to be
typed rather than partial: unchanged here, and the refusal this change removes
was a *false* one — it described well-formed markup as malformed. ADR 0006
requires preservation by default: the mark is a byte the producer wrote, so it
is carried through the model rather than dropped, which is why the fix adds the
prefix carry instead of simply discarding three bytes. Proposed ADRs 0030 and
0031 are not cited.

**Untouched contracts.** `Package::open`, `Package::save`, `Package::document`,
the source-backed routes, the managed document transaction, the OPC publication
plan and every relationship, content-type and signature rule are unchanged; the
census exercises nine of them per fixture and they are byte-identical.

## Measured

Deterministic counts only. No timing was run and none is required: the changed
code is a three-byte prefix comparison on a path that already parses the whole
part, and it is not shared with the reader.

### The corpus

`test-data/` holds **63** DOCX-family fixtures (62 `.docx`, one `.dotx`) — not
the hundreds the brief expected; `.docm` and `.dotm` are absent. Of the 63,
**exactly one** — `alt-chunk-header.docx` — has a UTF-8 byte order mark on
`word/document.xml` (`census/bom-census.txt`). That is why 0638 saw one fixture:
it is the only marked one in the repository. The corpus is dominated by
LibreOffice and POI output, which does not mark its parts; Word does.

### The admission census, before the change

Ten aspects per fixture, all through documented entry points
(`census/census-before.txt`, 630 lines):

| Aspect | ok | refused | skipped |
| --- | ---: | ---: | ---: |
| `Package::open` | 55 | 8 | — |
| `Package::document()` | 55 | 0 | 8 |
| `Document::sections()` | 55 | 0 | 8 |
| `Package::document_mut()` | **53** | **2** | 8 |
| `save` with no edit | 55 | 0 | 8 |
| `document_mut()` then `save`, no edit | 53 | 2 | 8 |
| managed `edit_document` + publish + `save` | 54 | 1 | 8 |
| `document_mut().add_paragraph_with_text` + `save` | 50 | 5 | 8 |
| reopen the edited artifact | 50 | — | 13 |

The eight `open` refusals are all package-level and all in
`litchi-opc`/`litchi-ooxml-common` (orphaned core properties, OPC M4.2–M4.5 core
property rules, two core-properties parts, derived part names). **None of them
is a DOCX-editor refusal, and there is no fixture the editor admits that the
reader refuses** — the asymmetry runs one way only.

The two editor refusals on reader-admitted fixtures are:

1. `alt-chunk-header.docx` — `invalid DOCX XML: syntax error: tag not closed`
2. `strict.docx` — `invalid DOCX format: invalid page width value '612pt'`

The three extra `edit_save` refusals are the three signed fixtures under
`poi/test-data/xmldsign/`, refused by the OPC signature-edit policy, not by the
editor.

### The minimal constructed witness

Bisecting `alt-chunk-header.docx`'s own `word/document.xml` (retained verbatim
as `witness/alt-chunk-header-document.xml`, sha256 `f573a236…40ee`, 3,476 bytes)
separates two independent triggers. Removing the three BOM bytes and nothing
else changes the refusal from the syntax error to a *structural* one; removing
the `w:altChunk` element and nothing else makes it parse.

The smallest input that reproduces the reported failure is five elements long:

```text
EF BB BF  <w:document xmlns:w="…/wordprocessingml/2006/main"
                     xmlns:r="…/officeDocument/2006/relationships"><w:body>
            <w:p><w:r><w:t>start</w:t></w:r></w:p></w:body></w:document>
```

— a marked document with one paragraph (`witness/bom-w1-paragraph.xml`). Before:
*tag not closed*. After: accepted. Its unmarked twin
(`witness/w1-paragraph.xml`) is accepted on both legs. A marked document whose
body holds **only** an empty `w:altChunk` (`witness/bom-w5-altchunk-only.xml`)
is accepted on both legs, because an empty element has no closing tag for the
shift to truncate. All fourteen witnesses and both legs' outcomes are in
`witness/witness-outcomes.txt`.

Two constructed package pairs isolate the mark at package level
(`witness/build-variants.py` rebuilds them from tracked fixtures): `a0/a1` from
`ooxml/docx/documentProperties.docx`, whose main document carries no
markup-compatibility markup, and `b0/b1` from `ooxml/docx/Hyperlink.docx`, whose
main document does. The marked member of each pair was refused before and is
admitted after; the unmarked member is unchanged. Their different pre-fix
refusal text is the same defect seen through different whitespace: where the
fixture is pretty-printed and the three-byte shift lands in indentation, the
scan survives and the truncated preserved range fails later; where the markup is
dense, the shift places a markup-compatibility marker comment inside the
`<w:body>` tag and the error names the wreckage —
`invalid XML QName 'w:bo<!--litchi-mce-active-…'`.

### Value identity

`diff census-before.txt census-after.txt` over 630 lines and 63 real fixtures
yields **three changed lines, all on `alt-chunk-header.docx`, all refusal text**:

```text
-  document_mut refused:invalid DOCX XML: syntax error: tag not closed: `>` not found before end of input
+  document_mut refused:invalid DOCX format: body-final section properties are not the final body child
```

and the same substitution on `touch_save` and `edit_save`. Every other line is
identical, which includes **every published SHA-256**: 55 no-op saves, 53
editor-touched saves, 54 managed-route saves and 50 edited saves, all
byte-identical between legs, and the reader's paragraph count and text length on
all 55 opened fixtures.

**The fixture that started this is still refused.** The syntax error was hiding
a structural one: `alt-chunk-header.docx` writes `w:sectPr` and then
`w:altChunk` as the last two children of `w:body`, and the editor requires the
body-final section properties to be the final body child. Fixing the offsets
does not admit it; it makes the refusal say what is actually wrong.

That diagnosis is corroborated from outside this change. Three sites state the
same rule independently — `writer/doc/package.rs:746`,
`document/transaction.rs:5594` and `:5610`, and `section/inventory.rs:934` — and
on the **before** leg the managed transaction was already refusing
`alt-chunk-header.docx` with exactly `body-final section properties are not the
final body child` while the editor reported a syntax error on the same bytes.
The after leg makes the two routes agree.

## Correctness evidence

**Three tests** in `crates/litchi-docx/src/writer/doc/tests.rs`, over a
pretty-printed body carrying a paragraph, a table, an `altChunk` and a `sectPr`:

* `a_byte_order_mark_shifts_no_preserved_body_range` — the marked and unmarked
  documents produce the same paragraph, table and anchor counts, the preserved
  paragraph survives intact with its attributes, and the marked document's
  `to_xml()` is exactly the mark followed by the unmarked document's.
* `a_byte_order_mark_survives_one_appended_paragraph` — after
  `add_paragraph_with_text`, the output still starts with the mark, contains it
  exactly once, retains the preserved paragraph and carries the appended text.
* `body_final_section_properties_must_remain_the_last_body_child` — the
  out-of-order body refuses with the same typed message with and without the
  mark, freezing both the contract and the fact that the mark no longer decides
  which refusal a caller gets.

**Differential corpus check.** The retained probe
(`probe/src/main.rs`, built once against the shared read-only checkout of
`c7326f680` and once against this branch, both `--release`) drives ten aspects
over 63 fixtures and four constructed packages on each leg. Results above; raw
output and diffs under `census/`.

**Gates**, all in the worktree, all passing, tails in `gates.txt`:
`cargo fmt --all --check`; `cargo clippy -p litchi-docx --all-targets` (deny);
`cargo test -p litchi-docx` (52 suites, 1,468 passed, 0 failed);
`cargo doc -p litchi-docx --no-deps` (deny); `cargo test -p litchi --features
docx,xlsx,pptx,xls` (26 suites, 266 passed); `cargo test` in
`tools/perf-baseline` (19 suites, 531 passed) and in `tools/native-resave`, the
two tool packages that link this crate and are reachable from no per-crate gate.
No pre-existing failure was encountered. One piece of **pre-existing drift** was
found and deliberately not committed: the `tools/native-resave` gate rewrites
that tool's `Cargo.lock` to add `xml-minifier` under `litchi-docx`, a dependency
`crates/litchi-docx/Cargo.toml:43` already declares on the untouched base
checkout; this change edits no manifest, so the lockfile was stale before it and
is left as the base has it.

## Validation preserved

Nothing in `validation.rs`, `sanitize.rs` or the OPC publication audit changed.
`validate_section_placement` and `final_section_properties` still run on every
`from_xml`, on the same elements, in the same order — they now receive the
ranges they were always meant to receive. The alternative-format scan's byte
ceiling and the shared element scanner's depth and node ceilings are untouched;
no validation path mutates its input.

## Limitations

**Not claimed.** No performance claim, no timing, no allocation or instruction
count. The change is not measured as faster or slower and is not expected to be
either; it adds one prefix comparison per `from_xml`.

**Four questions are frozen here, not resolved.**

1. **The body-final `w:sectPr` placement rule, in `litchi-docx`.**
   `alt-chunk-header.docx` places `w:altChunk` after the body-final `w:sectPr`.
   The editor refuses; the reader admits the file and reads its text. ECMA-376's
   `CT_Body` is `EG_BlockLevelElts*` followed by an optional `w:sectPr`, so the
   fixture is schema-invalid and the editor's refusal is defensible — but the
   two routes disagree, and the reader is the lenient one. Admitting it would
   mean modelling a body child after the final section properties, which changes
   what `content_insertion_index` means and where an appended paragraph lands.
   Witnesses: `witness/w3-sect-then-altchunk.xml` (refused) against
   `witness/w4-altchunk-then-sect.xml` (accepted), identical but for the order
   of the last two children.

2. **ST_PositiveUniversalMeasure in the writer's section parse, in
   `litchi-docx`.** `strict.docx` is an ISO Strict document
   (`xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"`) whose `w:pgSz`
   is `w:w="612pt" w:h="792pt"`. `crates/litchi-docx/src/writer/section/codec/xml.rs:409`
   parses that attribute with `parse_u32`, which admits only
   `ST_UnsignedDecimalNumber`, so `document_mut()` refuses with *invalid page
   width value '612pt'*. The reader opens the file, reads its five paragraphs
   and its one section, and the managed edit route publishes it. This is a
   feature gap in the mutable section model, not an offset defect; widening it
   changes which values are accepted and how they round-trip, so it belongs in
   its own record. This change neither widens nor narrows it — the census line
   is identical on both legs.

3. **The authored-XML publication audit's own byte order mark handling, in
   `xml-minifier` and reached through `litchi-opc`.** This is out of this
   change's scope by the brief and is **reported, not touched**. With the fix in
   place, a marked main document can be opened for editing and saved unedited,
   but a *regenerated* marked part is refused at publication:

   ```text
   I/O error: Failed to save package: XML publication rejected for '/word/document.xml':
   malformed XML at byte 0: invalid XML declaration boundary
   ```

   The cause is the same defect class in a different crate:
   `crates/xml-minifier/src/audit.rs:1319` slices `input.get(start..end)` with
   positions taken from `Reader::buffer_position`, so on a marked buffer the
   declaration event's raw bytes are `BOM + "<?xml …"` minus three trailing
   bytes and `check_declaration` (line 1542) rejects the boundary. That crate's
   *streaming* auditor already handles the mark explicitly (`saw_bom`,
   `bom_raw`, `bom_history` in `verify_reader_with_policy`); only the slice
   auditor does not. A pristine passthrough member is not audited, which is why
   the same bytes republish unchanged on the no-op route. Witness:
   `census/variants-after.txt`, the `a1`/`b1` `edit_save` lines.

4. **The managed document transaction has the same defect, in `litchi-docx`.**
   On a marked main document with a dense body, `edit_document` +
   `insert_paragraph` refuses identically on both legs with
   `invalid DOCX XML: ill-formed document: expected '</w:bo<w:p>', but '</w:p>'
   was found` — `crate::document::Snapshot` mixes quick-xml positions with its
   own buffer exactly as `DocumentBody` did, and `'w:bo'` is the same truncated
   `<w:body>` seen in the editor's markup-compatibility wreckage. On the
   pretty-printed real fixture the shift lands in indentation and the route
   reaches its section-placement check instead, which is why that fixture's
   `managed_edit` line is the placement refusal on both legs. It is **not**
   fixed here: that route carries source-backed splice reservations and
   exact-source byte proofs, so correcting it is a separate change with its own
   value-identity argument. Witness: the `managed_edit` lines of
   `census/variants-{before,after}.txt`, identical on both legs.

Consequently, **the end-to-end "edit a marked DOCX and save it" route still does
not work**, and this record does not claim that it does. What it claims is that
the first of the four obstacles is gone, that the diagnosis a caller now gets is
accurate, and that nothing else moved.

**Corpus reach.** One marked fixture in 63 is a weak corpus for a defect whose
real-world population is "documents written by Microsoft Word". The constructed
variants stand in for that population; they are synthetic.

## Retained evidence

[`results/change-0650/README.md`](results/change-0650/README.md) — the probe and
its manifest template, both census legs and their diffs, the BOM presence census
over all 63 fixtures, the fourteen witnesses with both legs' outcomes, the two
reproduction scripts, `follow-ups.md` (the four open questions above, each with
its site and witness), `decision.json` and `gates.txt`.
