# 0304 — No silent content loss on text sets

- Status: implemented on `uo2/text-loss`, cut from `feat/unified-ops-api`
  at `94f8047602`; trunk merged up to `e7fe30fb24` (records 0294, 0298, 0301).
- Date: 2026-10-16
- Input: record [0294](0294-acceptance-follow-ups.md) §2.4 (on
  `uo2/acceptance2`): a paragraph `text` set on ODT, OTH and ODM silently
  deletes the frames inside the paragraph, and setting the old text back
  does not restore them. Policy of records 0270/0271/0283: an edit either
  keeps the content it does not address or is refused with a typed error.
- Owned: the text-set paths of `litchi-odt` (body edit), `litchi-docx`
  (cell text), `litchi-pptx` (shape, paragraph and cell text), `litchi-ppt`
  (text body checks), `litchi-odp` (scene text changes) and the ODT and ODS
  adapters in `litchi-ops`; the `transaction.rs` `Leaf` boxing; the test
  `crates/litchi-ops/tests/text_set_preservation.rs`; the example
  `crates/litchi-ops/examples/text_set_bench.rs`.
- `performance_claim`: before/after medians of the touched set paths in §5
  (release, pinned core of a shared host, B A A B processes). No
  frozen-harness qualification.

## 1. The rule

A `text` set replaces the text of its container: character data, spaces,
tabs and line breaks, and the run wrappers that only carry formatting
(ODF `text:span`, DrawingML `a:r`, WordprocessingML `w:r` slots). The
formatting of the replaced text goes with it: that is what the set
addresses. Every other piece of content in the container is either **kept**
or the set is **refused** with a typed error, and a refused set leaves the
snapshot untouched (atomic in `edit` and `batch` too).

An element is kept when the format can put it back without inventing a
position; otherwise it is refused. Fields are always refused: keeping one
would add its text to the new text, and replacing it would lose it.

### 1.1 ODT, OTH, ODM: keep at the character offset

ODT, OTH and ODM share the ODT body edit (`litchi_odt::transaction::Edit::
set_text`, new module `crates/litchi-odt/src/body/kept.rs`):

- An element that adds nothing to `Node::text` is **kept**: frames and other
  drawing shapes, notes, annotations and their end markers, bookmarks,
  reference, index and change marks, soft page breaks, and fields that show
  no text. It moves out of any span and goes to the **character offset it
  had in the old text**, counted as `Node::text` reports it, with one line
  feed between the paragraphs of a block container. When the new text is
  shorter, the offset is clamped to its end. Elements at one offset keep
  their order.
- So setting the old text back puts every kept element back at its offset,
  and the source content bytes come back exactly when the text had no spans
  (tested on `writer-images-frames.odt`, `frames.oth`, `frames.odm`).
- **Refused** (`Error::Unsupported` → `unsupported_capability`, naming the
  element): an element that adds text of its own, such as a text field
  (`text:page-number`, `text:date`, `text:sequence`, `text:variable-set`), a
  hyperlink inside the node (`text:a`), `text:meta`, `text:ruby`, or a frame
  with text inside a text box. In a cell, list item, text box, note or
  annotation, a block child that is not a paragraph or heading (a nested
  table, list or section) is refused too. Bookmarks, change marks and soft
  page breaks at block level are kept as above.
- Setting the text of a hyperlink node itself (`/body/p[1]/hyperlink[1]`) is
  unchanged: the link is the node.

A plain paragraph (no `<` in its content) skips the plan entirely.

### 1.2 Other formats

Where a format has no offset model it can trust, it refuses rather than
keeps. Section 2 lists the per-format outcome.

## 2. Format × content kind × behaviour

"Kept" and "refused" were checked by a test in
`tests/text_set_preservation.rs` (one module per format) unless the cell
names the leaf test that covers it. **Fixed** marks a silent drop found and
removed by this record.

| Format | Container | Content kind | Behaviour |
|---|---|---|---|
| ODT/OTH/ODM | paragraph, heading, span, cell, list item, text box, note, annotation | frame, drawing shape, note, annotation, bookmark, mark, soft page break, empty field | **kept** at its character offset (**fixed**, was dropped) |
| ODT/OTH/ODM | same | text field, `text:a` inside, `text:meta`, `text:ruby` | **refused** `unsupported_capability` (**fixed**, was flattened or dropped) |
| ODT/OTH/ODM | cell, list item, text box, note | nested table, list, section | **refused** (**fixed**, was dropped) |
| ODT/OTH/ODM | same | spans and their formatting; several paragraphs | replaced (addressed); paragraphs merge into one |
| DOCX | paragraph, run (body and nested) | drawing, `w:pict`, comment, footnote and endnote references inside runs | kept (text slots are filled in order) |
| DOCX | paragraph | bookmark, comment range, `w:fldSimple`, hyperlink, content control | refused `unsupported_capability` (leaf `ComplexContent`) |
| DOCX | paragraph, run | complex field (`w:fldChar`/`w:instrText`) | refused (leaf `ComplexRun`) |
| DOCX | paragraph | a run with a reference and no text slot | refused (no text slot) |
| DOCX | table cell | picture, nested table, content control, field or bookmark in a later block | **refused** (**fixed**, was dropped) |
| DOCX | table cell | first paragraph the text rewrite refuses (bookmark, hyperlink, field) | **refused** (**fixed**, was rewritten as a bare paragraph) |
| DOCX | table cell | plain later paragraphs; an empty first paragraph | replaced (addressed) |
| PPTX | shape | `a:fld` in the text | **refused** `unsupported_capability` (**fixed**, the field's cached result was overwritten or blanked) |
| PPTX | paragraph | `a:fld` or other non-run markup between the first and last run | **refused** (**fixed**, became plain text) |
| PPTX | table cell | `a:fld` or other non-run markup in a paragraph | **refused** (**fixed**, was dropped) |
| PPTX | shape, paragraph, cell | line breaks, paragraphs, run formatting | replaced (addressed) |
| PPTX | run | `a:fld` as the target | sets the field's cached result (addressed) |
| PPTX | shape | markup-compatibility branches (`a14:m`) | refused (leaf, pre-existing) |
| XLSX | cell `value`/`text` | rich-text runs | replaced, the new value is plain (formatting of the addressed value) |
| XLSX | cell | comment, note drawing | kept (separate parts; tested byte-identical) |
| XLSB | cell | rich string runs | replaced (as XLSX) |
| XLSB | cell | comment, hyperlink | kept (comment part byte-identical, link target kept) |
| ODS | cell `value` | hyperlink (`text:a`) | **refused** `unsupported_capability` until `clear=true` removes it (**fixed**, was dropped); the same value is a no-op that keeps it |
| ODS | cell | spans, annotations, frames, several paragraphs | refused by the leaf for the whole row (pre-existing, typed) |
| DOC | paragraph | picture, field, note reference, other control characters | refused `unsupported_capability` (leaf dependency refusals) |
| DOC | paragraph | bookmarks | kept, shifted with the text (leaf; not retested here) |
| RTF | paragraph, run | any field, picture, bookmark, note, annotation in the body | refused `unsupported_capability` (whole document, pre-existing) |
| PPT | shape, paragraph | formatting runs, interactions, bookmarks, fields, several text atoms | length change refused (pre-existing) |
| PPT | shape, paragraph | metacharacter field (`*` for date, time, slide number, header, footer) | **refused** when the new text changes the character there (**fixed**, a same-length set overwrote it) |
| ODP | shape, cell, notes (story) | field, `text:a`, frame, note, annotation, bookmark in any paragraph | **refused** `unsupported_capability` (**fixed**, the story was rebuilt as plain text) |
| ODP | paragraph | same | **refused** (**fixed**, content was replaced) |
| ODP | span run | same, inside the span | **refused** (**fixed**) |
| ODP | field run | the field itself | refused (pre-existing) |
| ODP | shape, paragraph | spans and their formatting | replaced (addressed) |
| ODG | shape | anything but one plain text span (paragraphs, spans, links, fields) | refused `invalid_value` (leaf, pre-existing) |
| Pages | body paragraph | attachment anchored at the paragraph's start or end | kept (tested: the anchor survives) |
| Pages | body paragraph | attachment between text runs; line feed in the text | refused (leaf `InterleavedAnchors`, pre-existing) |
| Pages | section | attachment or footnote mark the new text would remove | refused (leaf `DependentContent`, tested) |
| Keynote | placeholder | object references (attachments, fields) the new text would remove | refused (leaf `DependentContent`, `litchi-keynote/tests/slide_text.rs`); no corpus deck has one, so the unified test covers plain placeholders only |
| Numbers | cell | rich text with object or field references | rewritten in place when uniquely owned, otherwise refused (leaf `RichText` dependency; `table_cell_edit.rs` unit tests) |

## 3. Code changes

Each crate's change is its own commit.

- `litchi-odt` (`59e515dc72`, `c8b3585325`): `body/kept.rs` plans what a
  text set keeps; `body/edit.rs` stores the plan with the staged text and
  splices the kept elements into the encoded text; `body/text.rs` exposes
  the reader's `collapse`/`skipped` so offsets match `Node::text`. The
  `Edit::set_text` rustdoc states the rule.
- `litchi-ops` ODT adapter (`63ddd26ae8`): the leaf's `Unsupported` refusal
  keeps the `unsupported_capability` code (other staging errors stay
  `invalid_value`).
- `litchi-docx` (`e1914e345a`, `59c6bf4b15`): `replace_cell_text_at` refuses
  a removed block that is not a paragraph of plain runs, and a first
  paragraph the text rewrite refuses unless it is plain (an empty one).
- `litchi-pptx` (`8cb9f4a8bc`): shape text slots carry a field flag; the
  paragraph and cell rewrites check the replaced children.
- `litchi-ppt` (`9494a35cbd`): the text body inspection records the
  metacharacter positions; one `check_replacement` replaces the four
  length checks.
- `litchi-odp` (`4ab2988697`): the scene records each paragraph's non-text
  inline content; the story, paragraph and run text changes refuse it.
- `litchi-ops` ODS adapter (`f56f7f59f8`): the hyperlink refusal, with a
  suggestion.

## 4. Lint: `Leaf` boxing

`cargo clippy -p litchi-ops --no-default-features --features
ppt,pptx,xlsb,odp,odg,oth,odm --all-targets -- -D warnings` failed
`large_enum_variant`: `Leaf::Pptx` held two `Editable` owners inline (216
bytes). `Leaf::Pptx` and `Leaf::Xlsx` now box their endpoints together with
the proof or patch they already boxed (`e6bf8c9377`), so a step allocates
as before, and a const assertion keeps `Leaf` at most two words in every
feature combination. The failing set and 17 other pairs and triples
(`docx,xlsx`, `pptx,xlsx`, `docx,pptx`, `pptx,odt`, `xlsx,odp`, `ppt,pptx`,
`xlsb,ods`, `doc,docx,rtf`, `xls,xlsx,xlsb`, `odt,ods,odp`,
`odg,odc,odi,odb`, `pages,keynote,numbers`, `pptx,keynote`, `xlsx,numbers`,
`docx,odt,pages`, `doc,xls,ppt`, `pptx,xlsb,odp`) are clean.

## 5. Measurements

`examples/text_set_bench.rs`, release, `taskset -c 9` (B A A B B A) then
`taskset -c 11` (B A A B), host load 8–11. "Before" is this tree with the
record's source changes reversed. Evidence and binary hashes:
[results/0304/](results/0304/) (`summary.txt` and ten JSON runs).

| Scenario | Before (µs) | After (µs) | After/before | Unified/leaf |
|---|---:|---:|---:|---:|
| ODT plain paragraph, 1,000 ¶ | 418.4 | 438.8 | 1.034–1.047 | 1.002–1.008 |
| ODT paragraph with a frame | 1021.5 | 1020.8 | 0.999–1.005 | 1.001–1.002 |
| PPTX shape | 101.5 | 99.6 | 0.977–0.979 | — |
| PPTX paragraph | 111.3 | 108.4 | 0.965–0.983 | — |
| PPTX table cell | 150.3 | 147.4 | 0.980–0.981 | — |
| PPT shape | 130.6 | 133.3 | 1.012–1.016 | — |
| DOCX table cell | 49.8 | 50.0 | 0.997–1.009 | — |
| ODP shape | 1265.2 | 1254.7 | 0.986–0.992 | — |

- Unified against leaf (ODT): within 1% — the adapter adds nothing; the
  check lives in the leaf.
- **ODT plain paragraph: +3.4–4.7% wall time, every run.** That path takes
  the plan's fast return (no `<` in the paragraph). Callgrind over both ODT
  scenarios and both legs counts +0.22% instructions (6.018 G → 6.032 G
  Ir), with memset, deflate and inflate dominating and each top function
  within 0.2%. The wall-time change is therefore not explained by work this
  record adds; it is under the 5% review trigger but is reported, not
  claimed as noise. Follow-up F2.
- No other touched path regresses beyond 1.6%.

## 6. Goldens

- `results/0289/fidelity.{md,tsv}` re-blessed after review: ODT
  set-then-revert rows that used to shrink the file are byte-exact (27 →
  32), keep more bytes, or are refused with the element named (4 → 8); OTH
  and ODM gain the two 0294 fixtures (byte-exact). No other format changes.
- `results/0289/crud-checklist.{md,tsv}`: unchanged (an intermediate
  version refused empty first cell paragraphs and changed the DOCX cell
  evidence; fixed in `59c6bf4b15`).
- `tests/golden/query_corpus/digests.json`: the 35 ODM digests changed
  because the ODM corpus gained `frames.odm` (§7, `d5cce163f6`); after the
  merge of trunk (record 0294 landed with the same fixtures) the file is
  trunk's and passes unchanged.
- After that merge the fidelity report was re-blessed again on trunk's
  version (`1caf73cec5`): the same ODT changes, and the `frames.oth`/
  `frames.odm` reverts are byte-exact instead of value-exact, which closes
  record 0294 §2.4.
- Capability and registry goldens (both feature sets), the OfficeCLI corpus
  golden and the capability delta table: unchanged.
- Python: `test_write_api.py` replaced the `*` of `datetime.ppt`'s date
  field with `x` — the overwrite this record refuses. It now edits the plain
  third shape of `text_shapes.ppt` (`01aeec5ab8`). Record 0294's Python CRUD
  mirror replays the Rust checklist evidence, which does not use that file.

## 7. Fixtures

`test-data/unified-ops/0294/frames.oth` and `frames.odm` were copied
byte-identical from `uo2/acceptance2` (record 0294's generator), with that
record's `PROVENANCE.md`; the trunk merge of record 0294 brought the same
blobs, so the merge was clean. Other
tests use tracked fixtures (`odf/corpus`, `odfdo`, `ooxml`, `ole`, `rtf`,
`poi`, `libreoffice-core`, `iwork`) or rewrite one member of one in the
test.

## 8. Gaps and follow-ups

1. **F1, ODT position rule limits.** Kept elements leave their spans
   (an `as-char` frame loses the span's character style), and a set whose
   new text is shorter than an element's offset clamps it to the end;
   setting the old text back then does not restore it.
2. **F2, ODT plain-paragraph wall time** (§5): +3–5% with no instruction
   increase; re-measure on an idle host or callgrind per scenario.
3. **F3, refusal codes differ by leaf**: ODG reports `invalid_value` for a
   text box it cannot replace; the others report `unsupported_capability`.
4. **F4, DOC `FloatingPictures.doc`**: paragraph 5 onwards fail a text set
   with `corrupt_file`, a typed but wrong code, found while probing (not a
   silent drop; not fixed here).
5. **F5, ODS rich cells** are refused for their whole row by the leaf; an
   edit to a plain cell sharing that row is refused too (pre-existing).
6. **F6, Keynote/Numbers** have no corpus file with inline objects in a
   text container; their refusals rest on leaf tests with synthetic
   storages.
7. ODT headers and footers are not text-settable through the unified API,
   so they were not audited.

## 9. Gates (exit 0 unless noted)

- `cargo fmt --all --check`.
- `cargo clippy -p litchi-odt -p litchi-docx -p litchi-pptx -p litchi-ppt -p litchi-odp -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings`.
- `cargo clippy -p litchi-ops --no-default-features [--features F] --all-targets -- -D warnings` for no features and F = odt, oth, odm, docx, pptx, ppt, odp, ods, xlsx, xlsb, doc, rtf, odg, pages, keynote, `ppt,pptx,xlsb,odp,odg,oth,odm`, plus the 17 combinations of §4.
- `cargo test --all-features` of `litchi-odt`, `litchi-docx`, `litchi-pptx`, `litchi-ppt`, `litchi-odp`; `litchi-ops` in chunks (`--lib`, `--doc`, then the integration tests in groups of 20); `litchi`. The first run of `query_corpus_differential` failed on the ODM digests (§6) and passes after the re-bless.
- `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features` for the six changed crates.
- `tools/check_crate_boundaries.py`, `tools/non_iwork_gate.py verify`, `RUSTC_BOOTSTRAP=1 tools/check_api_glossary.py` (39 profiles, 2,305 bans, 215 families), `tools/test_no_3rdparty_references.py`.
- `crates/pyo3-litchi/tests/run_tests.sh` (275 tests; 2 errors before `01aeec5ab8`, see §6).
- `cd tools/perf-baseline && cargo check --locked --all-targets`.
- `cargo test -p litchi-ops --test capabilities --test registry` with `--all-features` and with every feature but `encryption`: unchanged, not re-blessed.
- `cargo test -p litchi-cli --all-features --test officecli_corpus`; `tools/officecli_capability_delta.py`: no change.

## 10. Not claimed

- Rich-run formatting (XLSX/XLSB/ODS/ODT/ODP spans) is replaced with the
  text by design; this record does not preserve it.
- The audit covers `text` sets of containers. Other value sets (find and
  replace, template merge) use their own engines and were not re-audited.
