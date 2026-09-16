# Follow-ups opened by change 0650

Four questions this change witnessed and did not resolve. Each is stated with
the site, the witness that reproduces it, and what a resolution would have to
decide. None is acted on here.

## 1. The authored-XML publication audit reads a marked buffer three bytes low

**Site.** `crates/xml-minifier/src/audit.rs:1319` — `verify_with_policy` slices
`input.get(start..end)` with `start`/`end` taken from `Reader::buffer_position`,
which does not count the byte order mark quick-xml removed from the slice. The
declaration event's raw bytes are therefore `EF BB BF` followed by `<?xml …`
minus three trailing bytes, and `check_declaration` (`:1542`) rejects the
boundary. Reached from `litchi-opc`'s `xml_splice.rs:612`
(`xml_minifier::audit::verify_authored`), which raises
`OpcError::XmlPublication`.

**Witness.** `census/variants-after.txt`, the `a1-nomc-bom.docx` and
`b1-mc-bom.docx` `edit_save` lines:

```text
I/O error: Failed to save package: XML publication rejected for '/word/document.xml':
malformed XML at byte 0: invalid XML declaration boundary
```

**Why it is a defect rather than a rule.** The same crate's *streaming* auditor
already handles the mark explicitly — `verify_reader_with_policy` carries
`saw_bom`, `bom_raw` and `bom_history` — so the two auditors disagree about the
same input. A *pristine* passthrough member carrying the same bytes is not
audited and republishes unchanged, which is why change 0650's `noop_save` and
`touch_save` aspects succeed on marked inputs while `edit_save` does not.

**What a resolution must decide.** Whether a leading mark is part of the audited
document (in which case the slice auditor must offset its ranges, as the
streaming one does) or is forbidden in an authored part (in which case the
refusal is right but should be raised with its own message, and the pristine
passthrough path becomes the inconsistency). Out of change 0650's scope by its
brief; `litchi-opc` and `xml-minifier` were reported, not touched.

## 2. The managed document transaction reads a marked buffer three bytes low

**Site.** `crate::document::Snapshot` and the `Edit` layout scan in
`crates/litchi-docx/src/document/transaction.rs` mix quick-xml positions with
their own buffer exactly as `DocumentBody::from_xml` did before change 0650.

**Witness.** `census/variants-before.txt` and `census/variants-after.txt`, the
`a1`/`b1` `managed_edit` lines, **identical on both legs**:

```text
invalid DOCX XML: ill-formed document: expected `</w:bo<w:p>`, but `</w:p>` was found
```

`w:bo` is the same truncated `<w:body>` that appears in the editor's
markup-compatibility wreckage before the fix. On the pretty-printed real fixture
the three-byte shift lands in indentation and the route reaches its
section-placement check instead, which is why `alt-chunk-header.docx`'s
`managed_edit` line is the placement refusal on both legs.

**What a resolution must decide.** The same split change 0650 made is not
obviously safe here: this route carries source-backed splice reservations,
exact-source byte proofs (`Patch::target_for_exact_unmanaged_source`) and an
`Operation` inverse, so moving the buffer under those proofs needs its own
value-identity argument over the managed corpus. A separate record.

## 3. Body-final `w:sectPr` placement: the editor refuses, the reader admits

**Sites.** The rule is stated independently at
`crates/litchi-docx/src/writer/doc/package.rs:746`,
`crates/litchi-docx/src/document/transaction.rs:5594` and `:5610`, and
`crates/litchi-docx/src/section/inventory.rs:934`. The reader
(`Package::document()`, `Document::text()`, `paragraph_count()`, `sections()`)
states no such rule and admits the file.

**Witness.** `witness/w3-sect-then-altchunk.xml` (refused) against
`witness/w4-altchunk-then-sect.xml` (accepted) — identical but for the order of
the last two body children. The real fixture is
`test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk-header.docx`,
whose census lines show `reader ok paragraphs=1` beside `document_mut
refused:… body-final section properties are not the final body child`.

**What a resolution must decide.** ECMA-376's `CT_Body` is
`EG_BlockLevelElts*` followed by an optional `w:sectPr`, so the fixture is
schema-invalid and the editor's refusal is defensible; the question is whether
the *reader* should also refuse (narrowing what litchi can read from a real
LibreOffice-exercised file), or the editor should model a body child after the
final section properties (which changes what
`DocumentBody::content_insertion_index` means and therefore where an appended
paragraph lands). Change 0650 changes neither.

## 4. ISO Strict universal measures in the writer's section parse

**Site.** `crates/litchi-docx/src/writer/section/codec/xml.rs:409` parses
`w:pgSz/@w:w` with `parse_u32`, which admits only `ST_UnsignedDecimalNumber`.

**Witness.** `test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/strict.docx`
(`xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"`,
`<w:pgSz w:w="612pt" w:h="792pt"/>`). Its census lines, identical on both legs:

```text
strict.docx reader ok paragraphs=5 text=65bytes
strict.docx sections ok count=1
strict.docx managed_edit d715f7432c79…b7aa bytes=25661
strict.docx document_mut refused:invalid DOCX format: invalid page width value '612pt'
```

**What a resolution must decide.** `ST_TwipsMeasure` is a union of
`ST_UnsignedDecimalNumber` and `ST_PositiveUniversalMeasure`, so `612pt` is
schema-valid and the refusal is a feature gap rather than a defence. Admitting it
means choosing an internal representation and deciding what a round-trip
republishes — the original `612pt` or the converted twip count — which is an
output-byte question and therefore its own record. The reader's own
`section/codec.rs:1204` `parse_measurement` has the same restriction and would
need the same decision.
