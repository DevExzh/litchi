# 0341 — ADR 0008 rows in the CRUD checklist, the unreached cells, and the files that do not open (audit WP9)

- Status: implemented on `uo2/acceptance-adr8`, cut from `feat/unified-ops-api`
  at `c31198a5af`; trunk merged at `4282e69000` (record 0338), then the
  checklist and fidelity reports were regenerated.
- Date: 2026-10-17
- Input: the definition-of-done audit of 2026-10-17 (§4 layers 2 and 3,
  WP9), ADR 0008, design §14.
- Owned: `crates/litchi-ops/tests/{crud_matrix,corpus_fidelity,acceptance_fixtures}.rs`,
  `crates/pyo3-litchi/tests/test_crud_mirror.py`, `test-data/unified-ops/0294/`.
  Two small reader fixes, each in its own commit with a test, in
  `crates/litchi-ops/src/open/cfb.rs` (plus the two OLE detectors) and
  `crates/litchi-odp/src/authoring/flat.rs`. Two tests that the fixes made
  stale were updated (`odp_unified_read.rs`, the query-corpus digests). No
  shared dispatch file, `Cargo.toml` or iWork crate changed.
- `performance_claim`: none. The opener now checks for two stream names
  before the legacy host streams (the same `exists` lookups it already made
  later), and the flat ODP scanner admits a reference it used to refuse. No
  read or write path of an opened document changed, so nothing was benched.

## 1. ADR 0008 rows, per format

`crud_matrix.rs` now derives three rows per format from that format's
successful *changing* `set` cells (node sets before document-root sets).
The rows are golden-tested as `docs/unified-ops/results/0289/crud-adr0008.tsv`
(fixture, path, props, second path, second props, typed code, note). The
checklist summary gains three columns.

| Row | What the test does and requires |
|---|---|
| atomic failure | Writes the fixture's bytes to a file and opens it file-backed. It checks that item 0 (`set <path> <props>`) changes the document on its own. It then runs the atomic batch `[set <path>, set <missing sibling [9999]>]`. The batch must be refused with a typed code, and `Error::item() == Some(1)`. Afterwards the snapshot's bytes, the file on disk, a sentinel file under another format's extension and a sentinel file blocking the destination's parent must all be byte-identical. Both saves must be refused (typed), and the directory must hold no leftovers. |
| patch inverse | Runs `edit(set)` and requires `diagnostics.changed()`. `patch.apply(source)` must equal the commit's bytes, and `patch.inverse().apply(result)` must give the **exact source bytes**. The patch applied to its own result must be refused (typed: `invalid_input`, `PatchConflict`). Whether a byte-identical, separately opened copy takes the patch is reported but not required (see §5). |
| composition | Two sets on independent nodes, prepared on two threads from one snapshot. The test tries two instances of one element kind first (two paragraphs, two cells, two slides), then two element kinds, then two properties of one node. A then B and B then A must each hold both values, and each order must be repeatable. They must give the same bytes or, where the format writes new strings, links or records in edit order, the same document: text, node tree and both values. Each edit's patch on the other's result must be refused (typed), so neither edit can be silently dropped. |

Result over the merged tree: **63 of 63 rows ok, in 21 formats; 0 FAIL;
0 n/a.** `ADR_NOT_APPLICABLE` (format, row, reason) is empty. A listed row
that passes fails the test, and so does a row that fails without being listed.

- Atomic failure: `path_not_found` at item 1 in all 21 formats.
- Composition gives the same bytes in 15 formats: xlsx, pptx, xlsb, rtf, odt,
  ods, odp, odg, odc, odb, oth, odm, odf, pages, keynote. In 6 the bytes
  differ but the document is the same: docx and doc (hyperlink relationship
  order), xls (conditional-format records), ppt (records appended per edit),
  odi (`meta.xml`), numbers (string-table order).

**Pytest mirror.** `test_crud_mirror.py::Adr0008MirrorTest` replays all 63
rows through `litchi_py`:

- the atomic row in full: the code, `item == 1`, the File, the file on disk
  and both refused destinations;
- the composition row on a `ThreadPoolExecutor`: both values, repeatability,
  and the same bytes or the same text and node views, as the Rust row
  records;
- the patch-inverse row only up to its forward half. `Commit` has no
  `patch`, so the inverse and its conflicts cannot be reached from Python
  (gap G1).

## 2. Unreached cells: 13 → 0

| Cells | Fixture (`test-data/unified-ops/0294/`, provenance in `PROVENANCE.md`) |
|---|---|
| ppt `video` get/find/query/view (4) | `video.ppt` (43,520 B). It is Apache POI's `sound.ppt` (Apache-2.0) rewritten in place: its one 40-byte `ExWAVAudioEmbeddedContainer` becomes an `ExAviMovieContainer` of the same size, holding the same `ExMediaAtom` and an empty path atom. No offset moves. The PPT writer has no media frames. |
| keynote `connector` get/find/query/view/remove (5), `group` get/find/query/view (4) | `shapes.key` (456,779 B). It starts from `iwork/keynote/basic.key`, gets two rectangles and an image through `litchi_keynote`'s drawable edit, and then has its slide archive rewritten at the wire level. The first rectangle's Bézier path keeps only its move-to and first line-to, which is a straight line. The image becomes a `TSD.GroupArchive` that holds only its drawable and has no children. The leaf writes neither lines nor groups, and no corpus deck holds either. |

Both generators live in `acceptance_fixtures.rs`. Two generations are
byte-identical, and `LITCHI_BLESS=1` rewrites the files. Both files are
synthetic. The group is degenerate: it has no children, and the unified
model reads none.

## 3. Office-extension files that do not open: 40 → 37

Two reader fixes:

- **Word 2007 encrypted OOXML** (`Encrypted_MSO2007_abc.docx`; was
  `corrupt_file`, "OOXML-suffixed input does not have ZIP magic"). Word 2007
  writes a stub `WordDocument` stream next to `EncryptionInfo` and
  `EncryptedPackage`, and the CFB opener checked `WordDocument` first. An
  `EncryptedPackage` stream only occurs in encrypted OOXML, so the opener and
  the OLE detectors now check for it first. The file opens locked (Standard
  encryption) and unlocks with `abc`. Commit `e7884a6779`.
- **Flat ODP entity references** (`theme.fodp`, `alternativeText.fodp`; was
  "DOCTYPE and custom entities are not allowed"). quick-xml 0.41 reports
  `&apos;`, `&lt;` and `&#169;` as references, and the flat ODP scanner
  refused all of them. It now admits the five predefined entities and
  character references to XML 1.0 characters. Custom entities and DOCTYPEs
  are still refused. Both files open, and their no-op saves are byte-exact.
  Commit `291a05bb55`.

The other 37 files are classified in `corpus_fidelity.rs::NOT_OPENED`, with
the expected typed code. The test fails on any new refusal, on a changed
code, or when a listed file opens. The lenient class is opened with its
leniency in the test. Each file is listed in `fidelity.md`.

| Class | Files | Code |
|---|---|---|
| invalid input (spec-violating; strict reader) | 19: 9 RTF from LibreOffice tests (a bare `\rtf` ×2, `\brdrw` > 75, a negative page border, an odd `\datafield` hex count, a revision root outside its table, misordered shape properties, an unstarred `\shprslt`; plus `aggregate.fods`, 1,048,578 rows); `writer-minimal-nasty.odt` (`&nasty;`); 3 style fixtures with an empty `office:body` (2 fodg, 1 fodp); 2 POI OPC-conformance negatives (docx); 2 POI CFBs whose FAT sectors are not FATSECT (xls); `57456.xls` (SST total < unique); `SharedFormulaTest.xls` (XFCRC 72 ≠ 62) | `corrupt_file`, except the two CFBs, which are `unsupported_format` |
| fuzzed input | 6: LibreOffice `forcepoint-1/3/4/5.rtf`, `ww8/fail/redline-1.doc`, `binary-garbage-nested-groups.rtf` | `corrupt_file` / `unsupported_format` |
| legacy version | 4 DOC: Word 6.0 ×2, Word 95 ×2 | `unsupported_format` |
| invalid; opens leniently | 5 DOC with duplicate style names ([MS-DOC] 2.9); they open with `Leniency::TolerateStylesheetDefects` (checked) | `corrupt_file` |
| reader gap | 3: `stacked-column-chart.fods` (a chart inlined as a nested `office:document`, which is valid flat ODF); `test10001Encoding.rtf` and `test10007Encoding.rtf` (Mac code pages) | `corrupt_file` |

Total: 37 (doc 10, docx 2, fodg 2, fodp 1, fods 2, odt 1, rtf 15, xls 4).

## 4. Reports

`crud-checklist.{md,tsv}`, `crud-adr0008.tsv` and `fidelity.{md,tsv}` were
regenerated after the trunk merge.

- Checklist: 0 FAIL and 0 unreached claimed cells in every format, and the
  three ADR columns are ok everywhere.
- Fidelity: 997 exact, 0 differs, 0 error, 7 signed-exact, 15 locked (was 14).
- The query-corpus digests were re-blessed for flat ODP (7 opened, 1
  skipped; was 5 and 3) and Keynote (`shapes.key`). No other digest changed.

## 5. Gaps and follow-ups

- **G1 (bindings):** Python `Commit` exposes no patch, so the inverse and
  conflict halves of the patch row are Rust-only. Follow-up: `Commit.patch`
  with `inverse()` and `apply()`, named per ADR 0034.
- **G2 (documentation vs behaviour):** the `Patch` docs in `transaction.rs`
  say a separately opened file is a foreign snapshot "even if its document
  text or bytes match". The ODT/OTH/ODM and ODF object-tree handles
  (`same_snapshot` in `format/odt.rs` and `format/odf_tree.rs`) compare
  bytes instead, so a byte-identical reopened copy takes the patch. That
  covers odt, oth, odm, odg, odc, odi, odb and odf; the row notes say which.
  It is safe, because the patch is checked against the content, but the
  docs or the behaviour should be aligned (owner of `transaction.rs`).
- **G3 (litchi-ods):** flat ODS with an inline embedded-object document
  (`stacked-column-chart.fods`). The validator and the sheet, DDE and
  scenario parsers would all need to skip nested documents, which is not a
  cheap fix.
- **G4 (litchi-rtf):** `\ansicpg10001` and `\ansicpg10007` are refused as
  `corrupt_file`. They should be supported, or refused as unsupported.
- The composition row checks sequential composition plus stale-patch
  refusal at the unified layer. The unified API has no `join` (the XLSX leaf
  has one), so the "typed conflict" side is covered by stale-patch
  refusals, not by a join `ConflictSet`.
