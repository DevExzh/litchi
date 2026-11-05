# 0339 — Parity residue: corpus detail, capability rows for unlisted elements, SmartArt decisions (record 0332 F1/F2, audit WP12)

- Status: implemented on `uo2/parity-residue`, cut from `feat/unified-ops-api`
  at `e39d0a75a4` (record 0333 merge); trunk merged at `01d2284cc0`
  (record 0337). The merge changed no golden, delta, corpus or checklist
  output; the bench matrix was regenerated (`632b2b2628`).
- Date: 2026-10-17
- Input: record 0332 §5 and §7 (F1, the corpus detail bound; F2, the
  `unlisted` rows), record 0333's corpus section, and the definition-of-done
  audit of 2026-10-17 (WP12: XLSX SmartArt, PPT SmartArt, equations and
  masters/layouts).
- Owned: `crates/litchi-cli/tests/officecli_corpus*`, `tools/officecli_*.py`,
  `tools/officecli_dispositions.json`, and the capability gap rows in
  `crates/litchi-ops/src/format/*/spec.rs`. Each spec file changed in one
  hunk (the gap list) plus, where needed, one helper function. No shared
  dispatch file, leaf crate or iWork crate changed.
- `performance_claim`: none. The change adds static gap rows to
  `capabilities()` and changes test and tool code. No read or write path
  changed.

## 1. Corpus harness: skipped properties in full (record 0332 F1)

`crates/litchi-cli/tests/officecli_corpus.rs` bounded every outcome field at
48 characters. A warned call that skipped many properties ended its detail
with `+N`, and the report could not classify those names. In the v1 subset
they were 86 open lines, all "unverified". Record 0332 counted 83; the
record 0333 merge added three.

- The skipped list of a warned call, and the `;skipped:` part of a call
  that failed after skipping, is now recorded **in full**, in report order.
  Other details keep the 48-character bound.
- **Compact form:** a run of consecutive names that share a first dotted
  segment is written once, as `<head>.{<rest>,...}`. For example,
  `title.font,title.size,fill` is written `title.{font,size},fill`. The
  longest detail in the corpus is 230 characters.
- **Diff-friendly:** the grouping depends only on the report order, so a
  call's row changes only when its skipped set changes. The outcomes header
  documents the form.
- `tools/officecli_corpus.py` expands the groups (`expand_skipped`). It
  refuses a truncated `+N` list from an older outcomes file and asks for a
  re-record. The `(more)` pseudo-property and the `unverified` state are
  gone.
- The re-recorded `outcomes.tsv` differs from the base only in the detail
  column of skipped lists. Lines, verbs, results and targets are unchanged.

**How the 86 lines resolve:** all 86 are now **accounted**, as deferred-P1.
The skipped names the bound had hidden form 135 corpus groups. Each already
had a disposition through the capability delta or the format-wide corpus
entries, so `tools/officecli_dispositions.json` needed no new entry:

| Groups | Tier | Disposition |
|---:|---|---|
| 126 | P1 | deferred-P1 0332 (property breadth: chart title, plot-area and legend layout, data labels, point colours, ...) |
| 4 | P0 | planned WP16 (`presentation.category`, `extended.template`, `lastModifiedBy`, `revisionNumber`: root-level metadata) |
| 2 | P2 | deferred-P2 0332 (`revision.id`, media `trimEnd`) |
| 2 | P1 | planned 0335, now deferred-P1 (DOCX picture `hRelative`, `vRelative`; §4) |
| 1 | P1 | deferred-P1 0333 |

The full list is in `results/0339/newly-visible-skipped-properties.txt`.
The script that produced it is `results/0339/newly-visible-skipped-properties.py.txt`.
The lines that carry the four P0 groups are open because of WP16 already;
the bound was not what kept them open.

## 2. Capability rows for unlisted elements (record 0332 F2)

The delta table marked 42 OfficeCLI elements `unlisted`: classified, but
absent from `capabilities()`. Each format now lists them as typed element
gaps (`Subject::Element`) with a follow-up. The delta table shows **0
`unlisted` rows**.

| Format | Elements | Code | Follow-up |
|---|---|---|---|
| docx | `bookmarkEnd`, `permStart`, `ptab` | `unsupported_capability` | 0339-F2 (deferred-P1) |
| docx | `col` (OfficeCLI `column`) | `unsupported_capability` | WP4 (record 0335 F8) |
| docx | `fieldChar`, `instrText` | `unsupported_capability` | none: out of scope (record 0332): a complex field is one `field` node |
| doc, rtf | `chart`, `series`, `axis`, `bookmarkEnd`, `permStart`, `ptab`, `col` | `unsupported_capability` | record 0324 (DOC) or 0328 (RTF), deferred-P1 |
| doc, rtf | `fieldChar`, `instrText` | `unsupported_capability` | none: out of scope (record 0332) |
| odt | `bookmarkEnd`, `permStart`, `ptab`, `fieldChar`, `instrText` | `unsupported_path` (the navigator's `ABSENT` list) | none: not an OpenDocument text element |
| odt | `series`, `axis` | `unsupported_path` (`ABSENT`) | M6 ODT body index (record 0276), with the chart gap |
| docx, doc, rtf, odt, pptx, ppt, odp | `diagram` (OfficeCLI's Mermaid renderer) | `unsupported_path` | none: out of scope (record 0332, design §10) |
| xlsx, xlsb | `detectedtable` | `unsupported_capability` | 0339-F5 (deferred-P1) |
| xls, ods | `detectedtable` | `unsupported_capability` | WP10 or WP11 (record 0339 F5) |

- Each code is the one a path naming the element returns today, where the
  vocabulary knows the name. The Word vocabulary has no `diagram` and the
  workbook vocabulary has no `detectedtable`, so such a path is
  `invalid_path`; the rows use the closest gap code.
- The DOCX, DOC and RTF rows use a new `unlisted_gap` or `word_family_gap`
  helper. The other formats reuse their existing gap helpers. ODM and OTH
  inherit the ODT rows, and the encryption goldens follow DOCX and XLSX.
- **Dispositions:** none went stale. Each row was classified before, either
  by a format-wide entry (`docx-breadth`, `doc-family`, ...) or by a named
  one (`mermaid-diagram`, `complex-field-parts`, `docx-p0-residue`). Those
  entries still classify the rows, which are now `gap` instead of
  `unlisted`. The follow-up gives the current id, so `--check` finds no
  milestone-only follow-up.
- Record 0330's consistency test
  (`element_operations_agree_with_property_flags`) stays green: the new
  rows are gaps without elements.

### Tool fix: an element's own name wins over its path words

The delta tool looked a Litchi element up by an OfficeCLI element's names
in sorted order, and those names include the words of its paths. So
OfficeCLI's `slidemaster` (path `/slidemaster[N]/shape[K]`) matched Litchi's
`shape` before `slideMaster`, and `media` matched `audio`. This happened in
PPTX, PPT and ODP. The tool now tries the element's own name and file stem
first (`OcElement.lookup_order`), then the other names, sorted. A new unit
test covers it.

Effects:

- PPTX and ODP: `slidemaster`, `slidelayout` and `media` now compare against
  the right elements.
- **PPT:** `slidemaster` and `slidelayout` were reported `supported`; they
  are now the `gap` rows they are in `capabilities()` (deferred-P1 0323).
- Over the whole record, the delta has 7,373 → 7,376 gap rows; the
  matching fix accounts for the three new ones. P1 planned goes 467 → 445 and P1 deferred-P1 goes
  6,039 → 6,064: the 22 DOCX picture rows of §4 move to deferred-P1.

## 3. SmartArt, equations and masters/layouts (WP12)

**XLSX SmartArt: a deferred gap row, not an implementation.**
`capabilities()` now lists `smartart` as an XLSX element gap, with
follow-up 0339-F6 (deferred-P1): "a worksheet-drawing diagram model in
litchi-xlsx". The reasons:

1. **No leaf model.** `litchi-xlsx` reads and writes pictures, shapes and
   charts in worksheet drawings, but has no diagram model: no inventory of
   `dgm` relationships, no five-part authoring, no frame anchor (record
   0284 §7: "needs a worksheet-drawing diagram anchor in the leaf first").
   Typed SmartArt CRUD is design §13 "also in v1" because "models already
   exist" (the §12 principle). That holds for PPTX and DOCX, through
   `litchi-drawingml::diagram`. It does not hold for XLSX, so this record
   treats XLSX SmartArt as P1 breadth. The owner can overrule this.
2. **No parity pressure.** OfficeCLI has no SmartArt for any format, and no
   corpus line uses it, so no v1 line depends on it.
3. **Lane conflict.** An implementation would edit `litchi-xlsx` drawings and
   `format/xlsx*` at the same time as record 0338, which owns the XLSX P0
   residue.

The implementation remains tractable later. The PPTX path (records 0284 and
0335) shows the pieces: `diagram::authoring::Spec` stages the five parts,
and the XLSX chart insertion (`litchi-xlsx/src/chart/edit.rs`) already
writes a `graphicFrame` anchor into a worksheet drawing.

**PPT:** each item is a typed gap with a disposition. The follow-ups no
longer name milestones:

| Element | Code | Follow-up | Disposition |
|---|---|---|---|
| `smartart` | `unsupported_capability` | 0339-F3 (deferred-P1): PPT SmartArt and diagram groups. A legacy file keeps SmartArt and PowerPoint 97 diagrams as drawn OfficeArt groups, which read as group and shape nodes. | deferred-P1 0323 (ppt-family) |
| `equation` | `unsupported_capability` | 0323-F6: PPT Equation Editor objects (unchanged) | deferred-P1 0323 |
| `slideMaster` | `unsupported_capability` | 0339-F4 (deferred-P1): PPT main and title masters as nodes | deferred-P1 0323 (now also in the delta, §2) |
| `slideLayout` | `unsupported_capability` | 0339-F4 (deferred-P1): a PPT slide names a layout type in its slide atom, and there are no layout records to expose yet | deferred-P1 0323 |

PPT is outside the v1 boundary (`policy.p0Formats`), so all four are P1.
The CRUD checklist's 13 PPT `smartart` cells now cite 0339-F3, and its 26
master and layout cells cite 0339-F4 (`crud-checklist.tsv`).

## 4. Sweep of the open v1 lines

The v1 subset has 3,443 lines.

| | passed | accounted | open |
|---|---:|---:|---:|
| before (base `e39d0a75a4`) | 2,184 | 1,062 | **197** |
| after §1 (full skipped lists) | 2,184 | 1,148 | 111 |
| after §4 (dispositions) | 2,184 | 1,154 | **105** |

Open lines before and after, by blocker:

| Blocked by | Before | After |
|---|---:|---:|
| unverified (detail bound) | 86 | 0 |
| planned WP5 (`validate`, `raw-set`, `dump`, record 0336) | 57 | 57 |
| planned WP16 (XLSX residue, root metadata, record 0338) | 38 | 38 |
| planned WP4 | 9 | 9 |
| planned 0335 | 6 | 0 |
| planned 0333 | 1 | 1 |

I read every open line that is outside WP5 and WP16. None is a small fix
in the files this record owns:

- `word/pictures.sh:81, 106, 121, 135, 151, 165`: DOCX picture `crop`,
  floating (`anchor`, `wrap`, `behindText`, `h/vAlign`, `h/vRelative`,
  `h/vPosition`), `link` and `decorative`. These are P1 property breadth:
  the P0 picture keys (`src`, size, position, `alt`) are supported. They
  wait on a typed `wp:anchor` model in `litchi-docx` (record 0335 F1). New
  disposition **`docx-picture-breadth`: deferred-P1, ref 0335**, matching
  the PPTX picture breadth (`pptx-picture-table-breadth`, record 0333). The
  comment, series and axis rows of `docx-picture-comment-follow-ups` stay
  planned 0335. **6 lines accounted.**
- `word/tables.sh:189–194, 197`: a table-level `set` of `border.*` and a row
  `set` of `header` and `height.exact`. DOCX `table` and `row` have no `set`
  (`unsupported_element`), and `litchi-docx` has no table- or
  row-properties edit (its formatting transaction writes `tblPr` only when it
  creates a table). That is not small. They stay **planned WP4** (record
  0335 F8).
- `word/paragraph-formatting.sh:153–154`: `add paragraph` with
  `listStyle=bullet|ordered` and `start=5`. These are P0 (numbering basics):
  they create numbering definitions on the fly. They stay **planned WP4**.
- `ppt/tables/tables-nested.sh:86`: a table-level `set` of the six style
  flags (`firstRow` … `bandedCols`). `litchi-pptx` reads `a:tblPr` but has
  no table-properties write. It stays **planned 0333** (F5).

The table and row lines name only P1 breadth (borders, row height), but they
count as P0. The refusal is element-level, and the outcome records no
property for it, so the corpus group takes its element's tier. Follow-up
0339-F1 covers this.

## 5. Follow-ups

The ids below are cited by `capabilities()` and the dispositions.

| Id | What | Next step |
|---|---|---|
| 0339-F1 | The outcome of an element-level refusal (`unsupported_element`) names no property, so its corpus group takes the element's tier (P0) even when every property the call names is breadth | record the call's `--prop` names in the detail of an element-level refusal, and tier the group by them |
| 0339-F2 | Word-family markers: DOCX/DOC/RTF `bookmarkEnd`, `permStart`, `ptab` (deferred-P1) | standalone bookmark-end and permission-range nodes over the leaf markers; positional tabs as run children |
| 0339-F3 | PPT SmartArt (deferred-P1) | recognise SmartArt and diagram groups in the OfficeArt tree and project them as a read-only `smartart` |
| 0339-F4 | PPT masters and layouts (deferred-P1) | main and title masters as `slideMaster` nodes; layouts from the per-layout masters PowerPoint 2007 and later write |
| 0339-F5 | `detectedtable` (deferred-P1 in XLSX/XLSB; WP10 for XLS, WP11 for ODS) | a header-row block heuristic in table queries (OfficeCLI's synthetic node) |
| 0339-F6 | XLSX SmartArt (deferred-P1) | a worksheet-drawing diagram model in `litchi-xlsx` (read inventory, five-part authoring through `diagram::authoring::Spec`, a `graphicFrame` anchor), then the ops element and node CRUD as in record 0284 |

## 6. Verification

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-cli -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features f] --all-targets -- -D warnings`, f = none, doc, docx, odp, ods, odt, ppt, pptx, rtf, xls, xlsb, xlsx | 0 each |
| `cargo test -p litchi-ops --all-features` (`--lib` 929, `--doc`, then four `--test` chunks; 2,543 tests) | 0 |
| `cargo test -p litchi --all-features` (234) | 0 |
| `cargo test -p litchi-cli --all-features` (33) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-ops -p litchi-cli` | 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --test capabilities --test registry`, with `--all-features` and with every feature except `encryption` | 0, 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test crud_matrix --test corpus_fidelity` | 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test officecli_corpus` | 0 |
| `python3 tools/officecli_capability_delta.py --officecli …/3rdparty/OfficeCLI`; `--check` | 0, 0 |
| `python3 tools/officecli_corpus.py`; `--check` | 0, 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate` | 0 |
| `python3 tools/unified_bench_matrix.py --check`; `tools/check_feature_matrix.py` | 0, 0 |
| `python3 tools/check_crate_boundaries.py`; `tools/non_iwork_gate.py verify` | 0, 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (286) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |

Re-blessed: the capability and registry goldens, in both feature sets; the
delta table; the corpus outcomes and report; and the CRUD checklist TSV.
