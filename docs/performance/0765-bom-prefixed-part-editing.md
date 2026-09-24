# 0765: a part that begins with a UTF-8 byte-order mark reads, edits and saves like the same part without one — every quick-xml position in the OOXML crates now goes through one `ReaderOrigin`

Status: retained, correctness fix. `performance_claim: none` — the control
timings and instruction counts below are reported to show the fix costs nothing
material, not registered as a claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded and untouched.

Base `1d1044e3ac`; branch `perf/0765-bom-prefixed-part-editing`; production
commit `2a2c4a9791`, test follow-up `67edd9b7ed` (the DOCX differential also
reads the marked `word/styles.xml`). The coordinator's task: find every place in the OOXML
crates that converts quick-xml positions to byte offsets, fix each with one
shared helper, keep marks wherever a part is preserved, document any public span
whose meaning changes, test every format marked against unmarked, decide the
0744 lane and the 0747/0754 proofs, and show the benign path does not regress.

## The defect

quick-xml 0.41, built without its `encoding` feature as this workspace builds
it, removes one leading UTF-8 byte-order mark (`EF BB BF`) from its input before
its first event and does not count it: `buffer_position()` and
`error_position()` of a reader over a marked input are three bytes short of the
byte offsets in that input. Change [0677](0677-xml-publication-bom-offsets.md)
fixed this in the publication audit, [0650](0650-docx-editor-byte-order-mark-admission.md)
and [0670](0670-docx-parser-residues.md) in two DOCX paths, and a few PPTX and
DrawingML sites added their own `+3`. Every other site took reader positions as
byte offsets. Records [0744](0744-xlsx-eager-workbook-cell-path.md) and
[0755](0755-pptx-nested-text-run-panic.md) found the XLSX and PPTX cases.

Four read-only surveys classified all 430 position reads at base
([`sites/survey-summary.md`](results/change-0765/sites/survey-summary.md)).
Of those reads, 162 spliced and 55 sliced original bytes at the shifted
offsets, 41 exposed shifted public spans, and 2 put them into error messages;
the rest were already correct, relative, on fragments that cannot carry a mark,
or in tests. The effects
are of three kinds, and the new tests confirm each at base:

* **Refused edits.** The shifted splice produces malformed XML, which the
  site's own read-back or the publication audit refuses. This covers XLSX cell
  edits on both routes, PPTX `set_shape_text(s)`, shape tags, layouts, DOCX
  settings patches, OPC relationship and content-type splices, and more.
* **Wrong reads.** Examples: DOCX `Document::paragraphs()` text of a marked main
  part ("tag not closed"), PPTX `Shape::span`/`Common::xml`, DrawingML ink
  spans, modern-comment and survey extension bytes.
* **Silent corruption.** In an indented part the three bytes before an element
  are whitespace. Splicing at the shifted span then drops them and leaves the
  replaced element's last three bytes (`me>`, `"/>`) as text. The output is
  well-formed, so neither the read-back nor the audit catches it. At base, a
  PPTX theme colour and font replacement on a marked, indented theme published
  `</a:clrScheme>me>` and `</a:fontScheme>me>` and saved successfully
  ([`base-failures/theme-silent-corruption.txt`](results/change-0765/base-failures/theme-silent-corruption.txt)).
  The surveys found the same shape in XLSX worksheet, catalog, view, page, filter
  and validation edits; PPTX tracks, placeholders, layouts, custom shows and the
  font list; and DOCX settings, hyperlink detachment and web settings.

Marks occur in real files. 17 members of 3 of the 336 repository OOXML fixtures
begin with one ([`census/`](results/change-0765/census/census.tsv)):

* Word's `alt-chunk-header.docx`: the main part, headers, footers, settings,
  custom properties, relationships and content types.
* LibreOffice's `tdf167689_xmlMaps_and_xmlColumnPr.xlsx`: relationships, content
  types and VML.
* A POI slide master.

## What was changed

* **`litchi-core`: `xml::ReaderOrigin`** (new, `crates/litchi-core/src/xml/origin.rs`).
  - `ReaderOrigin::of(input)` records how many bytes precede reader position
    zero: 3 for a complete leading UTF-8 mark, else 0. A second mark does not
    count, and neither do UTF-16 marks, because this build does not strip them.
  - `offset(position) -> Option<usize>` and its inverse
    `position(offset) -> Option<u64>` are checked conversions.
  - Unit tests and a doc test cover it, and
    `crates/litchi-opc/tests/reader_origin_contract.rs` pins the model to the
    linked quick-xml. The contract covers the slice reader, `NsReader`,
    buffered readers, error positions, a second mark and UTF-16 marks. It also
    covers a split mark: a buffered reader whose first fill holds fewer than
    three bytes does not strip the mark.
* **Every OOXML reader converts through the origin of its exact input.**
  - `crates/litchi-{opc,ooxml-common,drawingml,xlsx,pptx,docx,xlsb,crypto,ole-common,xldm,formula}`:
    364 reads now convert through the origin — the helper, or the local rule
    in the two crates without `litchi-core`
    ([`sites/sites-after.tsv`](results/change-0765/sites/sites-after.tsv)).
  - Helpers that took a reader now also take its origin, so the compiler found
    every caller.
  - Fragment readers, whose origin is always zero, are converted too. Several
    consumers compared their own positions with spans from another reader, so
    every position in a crate has to be in the same frame. For example, PPTX
    tag and placeholder code compares against scene spans, and OPC
    `OwnedXmlPart` compares against caller ranges.
* **Ad hoc corrections replaced by the helper.**
  - These sites carried their own `+3`: PPTX SVG lifecycle (3 sites), the XLSX
    drawing source, the DrawingML theme family and the DOCX layout scan and its
    oracle.
  - Sites that strip one mark before parsing (DOCX `DocumentBody::from_xml`'s
    fused scan and the theme family scan) add the parsed slice's own origin.
    Their offsets therefore stay exact even for a doubled mark, which used to
    shift them silently.
* **The XLSX lane (0744) accepts marked worksheets.** In
  `crates/litchi-xlsx/src/raw/worksheet/lane.rs`, `Entry::locate` converts
  through the input's origin instead of declining, and `skip_to` does the same
  for the spliced input. The edit scanner's tail `shift` adds the origin
  (`scan.rs`), and the source facts converted through it (`codec.rs`).
* **Compactors carry the mark.** `raw::compact::changed`/`changed_worksheet`
  (XLSX) and `opened::xml::compact_changed_slide_xml` (PPTX) now emit the input's
  mark before the compacted events, as the DOCX writer already did (0650). The
  bytes are reserved with the input, so this adds no allocation.
* **Crates without `litchi-core`.** `litchi-xldm` and `litchi-formula` apply the
  same three-byte rule locally, each with a comment naming the helper; adding a
  dependency edge was not warranted.
  - `litchi-xldm`'s two scanners had been half-corrected: only the first event
    was re-based. Table and relationship renames of a marked part were refused.
  - `litchi-formula`'s OMML error position was three bytes early.
* **Tests** (each fails or differs at base except where noted):
  - Differential suites that run one scenario on a package and on its twin whose
    XML members carry a mark, compact and indented. They require identical
    observations and outputs that differ only by marks.
    - `crates/litchi-opc/tests/byte_order_marked_parts.rs` (5): topology
      publication, `OwnedXmlPart` ranges, eager round trip, a local replacement
      through the 0747 proof, and a refusal at the physical offset.
    - `crates/litchi-xlsx/tests/byte_order_marked_parts.rs` (5).
    - `crates/litchi-pptx/tests/byte_order_marked_parts.rs` (8).
    - `crates/litchi-docx/tests/byte_order_marked_parts.rs` (6).
  - Also: `crates/litchi-drawingml/tests/theme.rs` (1 new), xldm (1), formula
    (1) and crypto labels (1, with a marked error offset).
  - The XLSX lane, parse, scan and compaction tests that pinned the decline now
    require admission and parity. The 0755 pin now requires a correct edit that
    keeps the mark.

## Breaking changes (0652 trade-off 1)

Public byte ranges now count a leading mark: they are raw-byte spans that slice
the bytes they were read from.

* `litchi_pptx::shape::Span` (`Common::span`, `Shape::span`), relative to
  `Scene::xml()`. The owner keeps its mark unless markup-compatibility
  processing rewrote it.
* `litchi_pptx::shape::text::scan_ranges` offsets.
* `litchi_ooxml_common::xml::scan_omml_formula_ranges` ranges.
* The tag ranges `litchi_opc::OwnedXmlPart` edits accept: callers must pass byte
  ranges of `bytes()`.
* `litchi_drawingml::ink::SourceSpan` and the other ink and action spans.
* The DOCX `alt::scan` keys, `drawing::SourceDrawing` byte ranges,
  `revision::conflict` spans and content-control `SourceOccurrence` spans.
* The XLSX form-control MCE provenance ranges.

Values change only for inputs that begin with a mark; before, the same values
were three bytes early and sliced the wrong bytes. The docs of `Span`,
`scan_ranges`, `scan_omml_formula_ranges`, `OwnedXmlPart` and `SourceSpan` say
so. No signature changes except the new `ReaderOrigin`; no durable format
changes.

## Which edited parts keep the mark

* **Kept on every route that derives a part from its source bytes.**
  - Splices keep it with the untouched prefix: source-backed and eager XLSX edits,
    PPTX opened and source-backed edits, the DOCX managed transaction and settings
    patches, and OPC manifest and `OwnedXmlPart` edits.
  - Compactors carry it: XLSX and PPTX (new here) and DOCX (0650, 0670).
  - Untouched members are copied byte for byte.
* **Dropped only from parts the library regenerates from a model.** The library's
  serializers never emit a mark. The differential suites observed exactly these
  cases ([`marks.txt`](results/change-0765/marks.txt)):
  - the DOCX mutable writer's `word/_rels/document.xml.rels`;
  - `[Content_Types].xml` and a master's relationships after `add_slide_layout`;
  - a PPTX tag part rewritten by `put_shape_tags`.

  The independent review found more regenerated parts that drop the mark, all
  rebuilt from the model and none a preserved part:
  - `put_shape_tags` adding a tag part also drops it from `[Content_Types].xml`
    and the slide's relationships (28 decks);
  - XLSX eager cell edits that remove the calculation chain drop it from
    `[Content_Types].xml` and `xl/_rels/workbook.xml.rels` (45 workbooks);
  - DOCX settings edits drop it from the settings, document relationships and
    content types;
  - edits to a signed deck drop it from `_rels/.rels` when the signature is
    stripped.

  Every save without an edit kept every mark.
* The rule behind keeping marks through compaction is 0650's: under ADR 0006's
  preservation default, the mark is a byte the producer wrote, not formatting.

## The 0744 lane, the 0747/0754 proofs and the 0754 scanner

* **Lane: accepts.** 50 generated marked bodies take the lane in all three
  passes, with layouts, stores and compacted bytes identical to the reader's.
  Two further checks:
  - a marked layout equals the same document behind three whitespace bytes
    (`byte_order_marked_worksheets_take_the_lane_with_document_offsets`);
  - the end-to-end edit runs the same reduced readback as its unmarked twin and
    publishes the twin's bytes behind the mark
    (`byte_order_marked_worksheets_edit_exactly_like_unmarked_ones`).
* **0747 window proof: already correct.** It audits physical offsets since 0677,
  and its own test covers a marked original. Through OPC, a marked local
  replacement now publishes the unmarked bytes behind the mark. A marked
  malformed replacement is refused with an empty sink, at the unmarked refusal's
  offset plus three.
* **0754.** The layout scanner already added the mark and now uses the helper,
  with its NsReader oracle changed in lock-step; the scan-differential tests
  pass. The publication proof is exercised by the DOCX semantic-edit
  differential: the marked edit publishes the unmarked bytes behind the mark.
* **A doubled mark** is character data before the root, as XML says. It is now
  located exactly, and the publication audit refuses it at physical byte 3. It
  is not refused on every route (review correction):
  - on the DOCX package routes, the semantic edit (`edit_document`,
    `replace_paragraph_text`, `publish_document_edit`) and the mutable append
    both succeed and save the part with one mark, because compaction drops the
    second (`document/transaction.rs:5766`, `writer/doc/model.rs:1706`). The base
    does the same on the semantic route;
  - an XLSX worksheet with a doubled mark now passes `commit()` and is refused at
    `to_bytes()` by the publication audit, where the base refused it at
    `commit()`.
  Nothing is corrupted. Refusing a doubled mark on every route is a follow-up.
* **UTF-16 parts** (pre-existing, not changed): a DOCX main part encoded as
  UTF-16 reports `paragraph_count() == Ok(0)` while `text()` returns an error.
  Every edit route refuses such parts with a typed error.

## Evidence

* **New tests at base.** The differential suites were run on `1d1044e3ac`
  with the base's code ([`base-failures/`](results/change-0765/base-failures/)):
  16 of their 22 tests fail — OPC 2/3, PPTX 6/8, DOCX 5/6 and XLSX 3/5 — and so
  does the new DrawingML theme test, whose compact case returned malformed XML
  as `Ok`. The six that pass at base read or save without editing, edit a part
  without a mark, or are refused alike on both twins. Two OPC tests (the
  replacement and refusal-offset pins) and the xldm, formula and crypto tests
  were added after that run; the latter three exercise rules the base does not
  have.
* **At `67edd9b7ed`.** Every suite of the touched crates and their in-scope
  dependents passes; see Verification.

### Control timings (benign path)

The three cases are unmarked corpora, so they measure only the helper's cost on
the common path.

Legs:
* Base harness `058b58ba…` against branch harness `71453ca1…`. Both were built
  with `CARGO_TARGET_DIR=<target> CARGO_BUILD_JOBS=6 cargo build --release
  --locked --offline --manifest-path tools/perf-baseline/Cargo.toml --bin
  litchi-perf-baseline`, from worktrees with equal path lengths, and copied to
  equal-length paths.
* ABBA, 4 rounds, 8 processes per case, core 20.
* Every process is wrapped in `perf stat -e instructions:u,cycles:u`. That counts
  the whole process, including untimed corpus construction.
* Outputs and corpora are identical in every process.

| case | shape | samples/process | before p50 ms | after p50 ms | paired change (95% CI) | instructions | cycles |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `xlsx_first_cell` | dense-wide | 60 | 7.851 | 7.772 | −1.69% [−3.54%, −0.40%] | +0.32% | −1.84% |
| `pptx_semantic_one_edit_save` | large | 40 | 45.510 | 46.214 | +1.62% [+0.37%, +3.26%] | +0.10% | +2.20% |
| `pptx_semantic_one_edit_save` (confirmation) | large | 80 | 45.670 | 45.987 | +0.62% [−0.37%, +1.28%] | +0.10% | +1.95% |
| `docx_semantic_one_edit_save` | large | 100 | 5.598 | 5.612 | +0.55% [+0.23%, +0.99%] | +0.11% | +0.42% |

No paired ratio exceeds 5%. The PPTX first run's +1.6% did not reproduce in the
confirmation run, whose interval contains zero. Instructions move by at most
0.32% for the whole process, so these time differences are within the layout
movement two separately linked binaries show on this host (0746). Raw reports,
scripts and analysis: [`abba/`](results/change-0765/abba/).

## What is not claimed

* No performance claim, and no claim that the fix speeds anything up.
* The tests exercise representative routes of every format and every
  silent-corruption shape the surveys found. They do not run every public edit
  entry point on a marked part. The reads outside the tests were converted by
  inspection, listed site by site in `sites-after.tsv`.
* The origin models slice readers, and buffered readers whose first fill holds
  the whole mark. The streaming readers in these crates (the DOCX tail-append
  `GuardedBufRead` and ooxml-common's `mce::stream` prefix reader) consume the
  mark themselves and are unchanged.
* UTF-16-marked XML is not stripped by this quick-xml build and is not addressed
  here.
* ODF crates have the same reader behaviour and are out of scope (deferred), as
  are the iWork crates.
* Merge conflicts are expected with change 0764's attribute-handling edits in the
  same files. They meet only at the position lines.

## Verification

Gates ran on `67edd9b7ed` with non-incremental builds
([`gates.txt`](results/change-0765/gates.txt)). Every step below exits 0 unless
stated:

* `cargo fmt --all --check`.
* `cargo check --all-targets` of the twelve touched crates and
  `litchi-spreadsheet-drawing`, `litchi-doc`, `litchi-xls` and `litchi-ppt`.
* The facade with `doc,docx,ppt,pptx,xls,xlsx,xlsb,odt`.
* `cargo clippy --lib --no-deps -D warnings` of the touched crates.
* `cargo clippy --all-targets` per touched crate: clean except `litchi-pptx`
  (three `err_expect` lints in `opened/tests.rs`, pre-existing per 0755) and
  `litchi-xlsb` (`expect_used` in `comments/threaded/tests/mod.rs` and
  `shared_workbook/tests.rs`). This change does not touch those files.
* `cargo test` of the touched crates and dependents: 12,873 passed, 0 failed,
  89 ignored. The facade with the same features: 382 passed, 0 failed,
  7 ignored.
* `cargo doc -D warnings`.
* `check_crate_boundaries.py` and `check_perf_claims.py --mode structural`.
* `non_iwork_gate.py verify` fails as it does at base (litchi-xldm registration;
  a known failure of this base).

A first run on `2a2c4a9791` lost its test and doc steps when the shared disk
filled: the test log stops mid-line and the facade test and doc logs are empty.
Its other steps had the exit codes above. The target directory was then deleted
and every gate re-run from a clean build.

The repository's root `Cargo.lock` copy predates the spec-gap merge and lacks
`litchi-xldm`. `cargo metadata --offline` added that member and its dependency
edges without changing any package version. The file is gitignored and not
committed, and the harness uses its own tracked lock file.

## Cleanup

See [`results/change-0765/cleanup.json`](results/change-0765/cleanup.json).

## Retained evidence

[`results/change-0765/README.md`](results/change-0765/README.md).
