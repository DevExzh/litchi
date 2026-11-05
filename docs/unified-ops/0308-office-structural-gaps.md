# 0308 — Remaining non-iWork structural gaps (XLS, RTF, ODP, DOC, ODG)

- Status: implemented on `uo2/office-gaps`, cut from `feat/unified-ops-api`
  at `208e74e6ce`; trunk merged at `6c8d517eff` (record 0304) and
  `9940d9bfe0` (records 0306, 0307).
- Date: 2026-10-16
- Input: record [0293](0293-producer-structure.md) §4 gaps 3 (XLS) and 4
  (RTF); G-ODP-1 (records [0278](0278-odp-unified-adapter.md),
  [0288](0288-odf-follow-ups.md), [0292](0292-odf-verbs-and-flat-odf.md));
  record [0304](0304-no-silent-loss-on-text-sets.md) §8 F3 (ODG) and F4
  (DOC).
- Owned: `litchi-xls` (row/column shifts, Formula flags), `litchi-rtf`
  (producer paragraph splice), `litchi-odp` (scene styles and character
  formatting), `litchi-doc` (body text refusal), the `litchi-odg` text-set
  error kind, and the ODP adapter in `litchi-ops`.
- `performance_claim`: none; §7 shows the touched paths.

## 1. Result

| # | gap | outcome |
|---|---|---|
| 1 | XLS shifts on producer workbooks | row insert+remove byte-exact **176 → 227** of 305 sheets, column **165 → 227** (§2) |
| 2 | RTF producer structure | remove first **60 → 110** of 147, add+remove **7 → 121**, move+back **25 → 55** of 74 (16 byte-exact, 39 value-exact) (§3) |
| 3 | G-ODP-1 character formatting | **implemented** on runs and paragraphs: bold, italic, underline, size, colour, font, through derived automatic styles (§4) |
| 4 | DOC `FloatingPictures.doc` text sets (F4) | **typed refusal** `Refusal::RangeBoundary` (a smart-tag end inside the paragraph), `unsupported_capability`, instead of `corrupt_file` (§5) |
| 5 | ODG refused text sets (F3) | **`unsupported_capability`** (leaf `Error::Unsupported`), was `invalid_value` (§5) |

The tallies come from the ignored census
(`cargo test -p litchi-ops --all-features --test structure_census -- --ignored --nocapture`,
`CENSUS_FORMAT=xls|rtf`), one attempt per sheet (XLS, up to three per
workbook) or per file (RTF), before on trunk `208e74e6ce`, after on this
branch. XLS round trips count only when byte-exact.

## 2. XLS (`litchi-xls`, `cell_values/structural.rs`)

| cause (0293 census) | before | now |
|---|---|---|
| CONDFMT/CF12, DV, ShrFmla, Array, Table, AutoFilter **anywhere in the workbook** refused every sheet | 45 sheets | records of **other** sheets are admitted when no XTI of this workbook covers the shifted tab (no 3-D reference can reach it); the shifted sheet's own are certified with it |
| ShrFmla in the shifted sheet | refused | range moves whole (a shift point inside it is refused); absolute refs shift; a relative `PtgRefN`/`PtgAreaN` component stays when every cell and the cell it names are on the same side of the shift point; the cells' `PtgExp` anchors move; 3-D refs refused |
| CONDFMT / DV in the shifted sheet | refused | bounds and `sqref` ranges shift **when every rule's formulas hold no cell reference**; DVal carries no coordinates. CF12/CondFmt12 and referencing formulas stay refused |
| no INDEX record | 14 sheets | with no INDEX and no DBCELL, the row table is regenerated in the same row blocks and written back without either |
| duplicate DIMENSIONS | 6 | an embedded chart substream's DIMENSIONS is skipped (top-level records only) |
| producer DIMENSIONS wider than the cells | not byte-exact | kept, shifted, while they cover the cells, when the source's are not exact |
| MULRK/MULBLANK under a column shift | 11 | the run moves whole when the shift point is outside it; a point inside is refused |
| hyperlink with a location | 6 | admitted when the location names another sheet, a defined name or another document |
| split OfficeArt | 15 | the worksheet's `MsoDrawing` payloads (and their `Continue`s) are read as one stream; anchors are patched in place |
| comment / picture `Obj` | refused | FtNts, FtCf and FtPioGrbit subrecords are coordinate-free |
| complete reader dropped sheets with reserved Formula bits | 4 sheets | `[MS-XLS]` 2.4.127: reserved bits "MUST be ignored"; now ignored (bytes kept) |

**ClientAnchor semantics corrected.** The shift read flag 0 as fixed and
shifted flag 3. `[MS-XLS]` `OfficeArtClientAnchorSheet` sets `fMove`/`fSize`
when the shape is *kept intact*; Apache POI's `AnchorType` (0 move and
resize, 2 move don't resize, 3 don't move or resize) and LibreOffice's
position/size-locked flags agree. Now: 0 shifts like cells; 2 moves whole
when the shift point is outside it; 3 (and the spec-invalid 1) is refused
unless the shift is after it. Seven sheets with locked comment boxes that
the old code "shifted" are now refused (`a position-locked drawing anchor
is not shiftable without row and column sizes`). The public
`drawing_metadata::AnchorBehavior` names (`Fixed` for 0, `MoveAndSize` for
3) carry the same inversion; changing that read API is follow-up F1.

**Reader strictness kept where the spec does not allow more** (all still
drop the worksheet from the complete reader, so its shifts are refused):
WINDOW2 page-break zoom outside 10–400 (4 workbooks), `fShrFmla` without a
leading `PtgExp` (5; there is an opt-in compatibility profile), CONDFMT
rule counts above 3, a ShrFmla not after its anchor.

Census after, row add+remove (305 sheets): 227 byte-exact, 2 not
byte-exact (the two 0293 value-exact files), 4 sheets without a grid, 4
`corrupt_file` on fonts the reader refuses, and 68 typed refusals: 30
sheets the complete reader does not publish (§2, strictness kept), 8
position-locked anchors, 3 RgbExtra formulas, 3 non-comment `Obj`s, 3 CF
or DV formulas with references, and single-digit causes (CF12, Array,
spanning 3-D references, internal hyperlinks, cell-level duplicates).
Columns: the same 227.

Removing an existing row or column keeps 0293's rule (empty and
dependency-free only): 5 and 1 of 305.

## 3. RTF (`litchi-rtf`, `edit/producer_paragraphs.rs`)

The 0293 seam cuts the body into root-level `\par` segments and splices
them; the candidate is reparsed and must read back with every paragraph's
text, paragraph and run formatting and every collection unchanged. It now
also:

- **inserts** a paragraph (`Operation::RestoreParagraph`, the unified
  `add`) as `{\pard\plain text\par}`: the group scopes the reset, so the
  blocks after it keep their state, and removing it restores the source
  bytes;
- **splits a formatting group that holds several paragraphs**
  (`{\f1 one\par two\par}` → `{\f1 one\par}{\f1 two\par}`), each later group
  replaying the group's control words before it; a list marker group stays
  at the root;
- **replays inherited root-level state**: when a paragraph's predecessor
  changes, the root control words since the last point where both a `\pard`
  and a `\plain` were seen are written before it;
- treats **table rows** (`\trowd`/`\intbl` … `\row`, also a `\row` closing a
  group) as blocks that keep their place among the paragraphs that do not
  move, always with the state before them replayed; every table must read
  back unchanged (Debug equality of the model);
- lets a **removed** paragraph take the fields, pictures, notes and objects
  its segment holds whole, but not one end of a bookmark or comment range,
  unknown control words or unmodelled (opaque) content;
- includes a `{\listtext}`/`{\pntext}` marker in the first paragraph's
  segment, and recognises a root group that opens with `\pard` as a
  paragraph.

The exact splice runs first, so byte-exact results stay byte-exact; the
split and the replay are value-exact, proved by the readback. Section
breaks (`\sect`) are still refused.

| attempt | before | after |
|---|---|---|
| remove first paragraph (147 files) | 60 | 110 |
| add + remove (147) | 7 | 121 |
| move + back (74): byte-exact / value-exact | 15 / 10 | 16 / 39 |
| swap + back (74) | 15 / 10 | 16 / 40 |

Largest remaining refusals (remove first): bodies whose paragraph count is
not the root-level `\par` count (10), opaque content (9) or unknown control
words (4) in the removed paragraph, groups that close several levels at
once (4), section breaks.

## 4. ODP character formatting (G-ODP-1)

`litchi-odp::scene`'s scan now records where `office:automatic-styles`
holds its styles; `Scene::automatic_styles(xml)` indexes them on demand
(`AutomaticStyles::{get, paragraph_text_properties, run_text_properties}`;
`AutomaticStyle`, `TextProperties`), so opening and reading a presentation
does not pay for them (a first version indexed them in the scan and cost
+38% on a 400-style scene, §7). The new
`Change::SetTextProperties { story, paragraph, run, properties:
TextPropertyEdit }` (named like the ODT edit):

- copies the node's automatic style under a fresh name (`T<n>`/`P<n>`),
  rewrites or adds its `style:text-properties`, and points the node at the
  copy; a node whose style is common (or absent) gets an automatic style
  whose parent it is. **No existing style changes**, so runs sharing a
  style keep it;
- wraps a plain-text run in a new `text:span` (folding a text set of the
  same run in the same plan);
- inserts the definitions before `</office:automatic-styles>`, expands an
  empty element, or creates one before `office:body`;
- is an exact no-op when the style already has the values; refuses link
  and field runs and a root that binds no `style`/`fo` namespace.

The adapter adds readable and settable `bold`, `italic`, `underline`,
`size`, `color`, `font` to runs and paragraphs (local properties of the
automatic style, the ODT convention); the handle indexes the styles once,
on the first such read. Coalesced batches track formatting
per story (`Footprint::Format`) so a format never shares a run with a
pending text write; before, a staged change without a footprint would have
been dropped after a barrier. Shapes, table cells and notes take no
character formatting (the `set shape character-format` gap now says so).

Tests: `litchi-odp` `scene::tests` (5 new), `litchi-ops`
`tests/odp_character_format.rs` (bytes outside the span tag and the new
style are unchanged; no-op; text and format in one call; coalesced batch
with inverse; refusals).

## 5. DOC and ODG refusal codes

- **DOC F4.** `FloatingPictures.doc` tags every style name as a smart tag;
  the `PlcfBklFactoid` (FIB table 117) end lies inside paragraphs 5 on. The
  replacement reached the revision editor's destructive-interaction guard,
  which reports a corrupt file. `body_text::Edit::replace_text` now checks
  the modeled CP tables first (after the exact no-op check) and returns
  `Refusal::RangeBoundary { fib_index }`, whose message names what the
  table positions ("smart tag"); the unified code is
  `unsupported_capability`. A tag ending at the paragraph's edge stays
  editable.
- **ODG F3.** `set_shape_text` on a shape whose text is not one plain span
  returns `Error::Unsupported` (packaged and flat), so the adapter's
  existing mapping reports `unsupported_capability`.

## 6. Goldens and fixtures

- `results/0289/fidelity.{md,tsv}` re-blessed after review on the merged
  tree: XLS structure byte-exact 58 → 80 (refused 59 → 37), revert
  byte-exact 100 → 103; RTF structure byte-exact 15 → 16, value-exact 10 →
  39, refused 49 → 18, one-way 0 → 1 (`tdf158830.rtf`: the move is
  admitted, its inverse is refused at a table row). A script compared every
  changed row: no file got a worse outcome.
- `results/0289/crud-checklist.tsv`: XLS read rows now cite
  `external-ref.xls` (the complete reader publishes it); ODP paragraph and
  run `set` evidence is a `bold` set.
- `tests/golden/query_corpus/digests.json`: 89 XLS digests changed (cells
  118,668 → 118,712) because the complete reader no longer drops the
  worksheets with reserved Formula bits; no other format changed.
- Capability and registry goldens (both feature sets) and
  `officecli-capability-delta.md`: ODP only (new run/paragraph properties,
  the run-format gap removed, the shape gap reworded).
- No new fixture; all tests use tracked `test-data/` files.

## 7. Measurements

`examples/office_gaps_bench.rs` (compiles against the base tree too) and
`examples/office_gaps_new_bench.rs`; release, `taskset -c 9` on a shared
32-core host (load 6–8), 15 rounds after 3 warm-ups, medians of per-call
means, paired legs ABBA, binaries alternated B A A B B A. "Before" is
trunk `6c8d517eff` with the example copied in. Evidence, binary hashes and
the build command: [results/0308/summary.txt](results/0308/summary.txt)
(`r2-*` are the final runs).

| scenario | before µs | after µs | after/before | unified/leaf after |
|---|---:|---:|---:|---:|
| XLS complete open, 20k cells | 5,791 | 5,643 | 0.974 | — |
| XLS complete open, FormulaEvalTestData | 1,854 | 1,812 | 0.977 | — |
| XLS row insert, 10k cells (native / unified) | 9,954 / 9,949 | 10,023 / 10,037 | 1.007 / 1.009 | 1.000–1.001 |
| RTF move, 1k plain paragraphs | 435 / 435 | 431 / 431 | 0.99 | 0.996–0.998 |
| RTF move, 1k producer paragraphs (exact splice) | 605 / 605 | 601 / 600 | 0.99 | 0.988–1.005 |
| ODP `Scene::parse`, 200 slides, 400 styles | 488 | 504 | 1.034 (see below) | — |
| DOC paragraph text set | 58.5 / 91.0 | 57.5 / 90.0 | 0.98 / 0.99 | 1.556–1.567 |

New paths, unified against leaf (two runs):

| scenario | leaf µs | unified µs | ratio |
|---|---:|---:|---:|
| ODP bold + size on one run, 200 slides | 1,815 / 1,822 | 1,824 / 1,821 | 1.005 / 1.000 |
| RTF producer insert, 1k paragraphs | 869 / 855 | 856 / 848 | 0.985 / 0.992 |
| RTF value-exact move (replayed state), 1k | 730 / 740 | 730 / 737 | 0.999 / 0.995 |
| XLS row insert, sheet without INDEX | 95.8 / 95.2 | 97.2 / 96.3 | 1.015 / 1.012 |

- **ODP scene.** The first version indexed every automatic style in the
  scan: **+38%** on this scene (r1). Styles are now indexed on demand
  (§4). The remaining wall-time difference was +3.4% in one sequence and
  +6.6% in another (before legs 459–495 µs across runs); callgrind over
  the whole scenario counts **3.615 G → 3.610 G instructions (−0.14%)**,
  so no added work explains it. Reported, not claimed as noise.
- **XLS row insert.** The first version re-read the source row table and
  the workbook's BoundSheets on every shift: +5.1% (r1). Both are now read
  once or only when needed: +0.7–0.9%.
- **DOC unified/leaf 1.56** is pre-existing (base 1.545–1.58): the unified
  set does the reader's paragraph resolution and the publication the leaf
  leg skips. Outside the single-op budget; unchanged by this record.
- The XLS reader's 2–3% gain is not claimed (the change is one mask).

## 8. Gaps and follow-ups

1. **F1 XLS `AnchorBehavior` naming** (read API): the variants carry the
   inverted fMove/fSize reading corrected in the shift (§2).
2. **XLS:** CF12/CondFmt12; CF and DV formulas with cell references
   (relative to the range's first cell); Array and Table; split shared
   formula ranges; position-locked anchors (need row/column sizes);
   reader strictness outside the spec (WINDOW2 zoom, `fShrFmla` without
   `PtgExp`, CONDFMT count, orphan ShrFmla), and the duplicate-Number and
   RgbExtra refusals.
3. **RTF:** section breaks; bodies whose paragraph count is not the
   root-level `\par` count; opaque or unknown content in a removed
   paragraph; the inverse of a move across a table row
   (`tdf158830.rtf`).
4. **ODP:** shape, table-cell and notes character formatting; link and
   field runs; reading formatting through common styles (`styles.xml`).
5. **DOC:** keep a smart tag across a text set (shift it with the text)
   instead of refusing.

## 9. Gates (exit 0)

Run on the branch after both trunk merges (`6c8d517eff`, `9940d9bfe0`);
cargo under the memory-capped scope shim after the 07:46 session loss.

- `cargo fmt --all --check`.
- `cargo clippy -p litchi-xls -p litchi-rtf -p litchi-odp -p litchi-doc -p litchi-odg -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` (and `-p litchi-ops -p litchi -p litchi-py` again after the second merge).
- `cargo clippy -p litchi-ops --no-default-features [--features F] --all-targets -- -D warnings` for no features and F = xls, rtf, odp, doc, odg, `doc,xls,rtf`, `odp,odg`.
- `cargo test --all-features` of `litchi-xls`, `litchi-rtf`, `litchi-odp`, `litchi-doc`, `litchi-odg`, `litchi`; `litchi-ops` `--lib` (915), `--doc` (207), and its 136 integration tests in seven groups of 20 (groups 3–6 rerun after the session loss with `RUST_TEST_THREADS=4`); after the second merge the record's tests, `capabilities`, `registry`, `query_corpus_differential`, `corpus_fidelity` and `crud_matrix` again.
- `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features` for the six changed crates.
- `tools/check_crate_boundaries.py`; `tools/non_iwork_gate.py verify`; `RUSTC_BOOTSTRAP=1 tools/check_api_glossary.py` (39 profiles, 2,305 bans, 221 families); `tools/test_no_3rdparty_references.py`; `tools/unified_bench_matrix.py --check`.
- `crates/pyo3-litchi/tests/run_tests.sh` (279 tests).
- `cd tools/perf-baseline && cargo check --locked --all-targets`.
- Goldens: capability and registry in both sets (re-blessed for ODP, §6); `officecli_capability_delta.py` regenerated; `cargo test -p litchi-cli --all-features --test officecli_corpus` and `tools/officecli_corpus.py --check` unchanged.

