# 0324 — DOC element gaps against the DOCX vocabulary (tier 2)

- Status: implemented on `uo2/doc-gaps`, cut from `feat/unified-ops-api`
  at `db96990ff6`; trunk merged at `a4a045b48a` (record 0322).
- Date: 2026-10-17
- Input: the DOC rows of `officecli-capability-delta.md`, which listed
  `styles`/`style`, `numbering`/`abstractNum`/`num`/`level`, `hyperlink`,
  `bookmark`, `field`/`toc`, `section`, `tab`, `textbox`, `shape`,
  `equation`, `formfield`, `revision`, `sdt` and `watermark` as gaps or
  missing. Records 0258 (DOC P0 surface) and 0319 (DOCX P0/P1 elements)
  are the starting point.
- Owned: `litchi-doc`; `litchi-ops/src/format/doc*`; the DOC rows of the
  capability and registry goldens, the CRUD checklist and the fidelity
  report; small hunks in `capabilities.rs`, `capabilities/tests.rs` and
  `error.rs`.
- `performance_claim`: none; §6 reports every new and touched path.

## 1. Result

DOC answers the DOCX paths and property names of record 0319.

| # | element | get / query | set | add | remove |
|---|---|---|---|---|---|
| P0 | `/styles`, `/styles/<id or name>` | ✅ `count`; `id`, `name`, `type`, `default`, `basedOn`, `next`, `linked`, the UI flags, `uiPriority`, `customStyle`, and the style's own character and paragraph properties | — | — | — |
| P0 | `/numbering`, `abstractNum[N]`, `num[N]`, `level[L]` | ✅ counts; definition `type`, `name`; instance `abstractNumId`, `startOverride.N`; level `start`, `format`, `lvlText`, `justification`, `suff`, `lvlRestart`, `isLgl`, `indent`, `hanging`, number font, size, colour, bold, italic | — | — | — |
| P0 | paragraph `style`/`styleId`/`styleName`, `numFmt`, `listStyle`, `start`, `numLevel`; run `rStyle` | ✅ | ✅ `style`/`styleId`/`styleName` (id or display name), `rStyle` | — | — |
| P0 | `/body/p[N]/hyperlink[M]` (`HYPERLINK` fields) | ✅ `url`, `anchor`, `tooltip`, `tgtFrame`, the first result run's formatting, `rStyle`, text | ✅ `url`, `text` | ✅ `url` or `anchor`, `tooltip`, `text`; at the end, the start, or beside a run | ✅ the whole field |
| P1 | `/bookmark[N]`, `/bookmark[@name=X]` | ✅ `name` | ✅ `name` (rename) | ✅ under the body (a new paragraph holding `text`) or a paragraph (over the first occurrence of `text`, or empty at its end) | ✅ (its text stays) |
| P1 | `/section[N]` | ✅ OfficeCLI's page-setup keys: `type`, page size and orientation, the seven margins, `columns`, `columnSpace`, `columns.separator`, `titlePage`, `pageNumFmt`, `pageStart`, line numbering, `direction`, `rtlGutter`, `vAlign` | ✅ `type`, page size, orientation, the margins, `columns`, `columnSpace`, `titlePage`, `vAlign` | — | — |
| P1 | `/field[N]`, `/toc[N]` | ✅ `fieldType`, `instruction`, `fldLock`, `evaluated`, result text; TOC `levels`, `hyperlinks`, `pageNumbers` | — | — | — |
| P1 | `p[N]/tab[M]` (tab stops) | ✅ `pos` (twips), `val`, `leader`, on any paragraph | ✅ body paragraphs | ✅ | ✅ |
| P1 | `/body/textbox[N]`, `/body/shape[N]` | ✅ text box text and extent; shape extent, `anchor.x`/`anchor.y`, `wrap`, `hRelative`, `vRelative`, `behindDoc` | — | — | — |

The remaining element gaps, each with its follow-up (§7), are `equation`,
`formfield`, `picture`, `ole`, `revision` and `watermark`; `sdt` is
`unsupported_path` (DOCX only).

OfficeCLI delta table, DOC against the DOCX vocabulary (base `db96990ff6`
→ this record):

| | supported | gap | missing |
|---|---:|---:|---:|
| elements | 15 → **29** | 11 → 5 | 18 → 10 |
| property reads | 39 → **156** | element gap 231 → 59 | 485 → 507 |
| property writes | supported 1 → **10**, partial 13 → 29 | 782 → 757 | |

Missing reads rise because properties of elements that were element gaps
(and counted there) are now compared key by key; the DOCX keys DOC has
no store for (theme fonts, `font.ascii`/`font.hAnsi`, East Asian and
character-unit indents, `lang.*`, the effective-value keys) stay missing.

CRUD checklist, DOC: 13 → 27 elements, claimed cells 86 → 188, all ok
(`results/0289/crud-checklist.{md,tsv}`).

## 2. Leaf surface (`litchi-doc`)

Each verb of the primary `doc::Edit` stages one MS-DOC structure through
the revision editor, appends the rewritten tables and publishes once.
Every other byte is kept; a patch inverts to the exact source bytes; a
value the target has is an exact no-op (the commit reports no change).

| Verb | Structure | Refusals |
|---|---|---|
| `set_paragraph_style(Position, istd)` | the paragraph's PAPX `istd`; the PAPX run is split so no other paragraph changes, and the PAPX FKPs are rebuilt | `StyleNotFound` (empty or non-paragraph slot), `TrackedText` (tracked mark or `sprmPPropRMark`), huge PAPX |
| `CharacterValue::Style(istd)` | `sprmCIstd` through `set_character_value`; index 10 clears it | `StyleNotFound` |
| `paragraph_tabs`, `set_paragraph_tabs` | one `sprmPChgTabsPapx` that clears every stop the paragraph shows (and every stop its old change deleted) and adds the new ones | `InvalidProperty` (±31680, duplicates, more than 64), `FormattingDependency` (direct `sprmPChgTabs`) |
| `insert_hyperlink(Position, offset, HyperlinkInsert)` | the field's marks (`fSpec`), instruction and result as new CHPX runs in one Unicode piece, three `PlcffldMom` markers (`flt` 0x58, `fHasSep`), every modeled PLCF shifted | `ComplexField` (inside another field), `TrackedText`, `PositionDependency` (unmodeled CP tables, ambiguous `Selsf`) |
| `remove_field(start_cp)` | the field's text and its markers (nested fields included) | `FieldNotFound`, `AnchoredContent`, `TrackedText` |
| `TextTarget::Field { start, part }` | a field's instruction or result as a text target of `replace_text` | `ComplexField` (a nested field in the part) and the text-replacement refusals |
| `bookmarks`, `add_bookmark`, `rename_bookmark`, `remove_bookmark` | `SttbfBkmk`, `PlcfBkf` (keeping each surviving bookmark's `bkc`) and `PlcfBkl` | `BookmarkName`, `BookmarkNotFound`; a `SttbfBkmk` with per-name extra data is refused (not modeled) |
| `set_section_values(Position, &[SectionValue])` | a new SEPX appended to `WordDocument`, the section's `SED.fcSepx` repointed | `SectionNotFound`, `InvalidProperty` |

`ChangeRef` gains `ParagraphStyle`, `ParagraphTabs`, `Field`, `Bookmark`
and `Section`; the durable patch envelope refuses them, as it refuses the
structural changes of record 0258. Three-way plans replay them.
Tests: `crates/litchi-doc/tests/doc_marks.rs` (6 tests, each reading the
published bytes back through the physical reader and inverting the patch).

## 3. Adapter (`litchi-ops`)

- `format/doc/parts.rs` (styles and numbering), `marks.rs` (hyperlinks,
  bookmarks, fields, tables of contents, sections, tab stops, text boxes,
  shapes) and `marks_write.rs` (their writes); new kinds in `spec.rs`,
  navigation arms in `nav.rs`, node items in `node.rs`.
- **Style ids.** DOC stores names only. A style's id is the rule Word
  applies when it saves DOC as DOCX: the name without spaces and
  punctuation, each word of a built-in style's name capitalized
  (`heading 1` → `Heading1`); a taken id gets the index appended. The
  ids are derived once per handle. `/styles/<step>` matches the id, the
  name or an alternate name, exactly.
- **Views stay as they were.** The new kinds are reached by path and
  query; the generic views of the root, the body and a paragraph list the
  structural kinds only (`views_list_only_the_structural_children`).
- **Native identities are raw**: `abstractNum`/`num` `id`, level `ilvl`.
- **Value projection** reads the parsed, detached reader without a source
  fence per key; path resolution and queries keep the fence.
- `set` checks for a mark path by the last step's name before resolving,
  so a paragraph or run `set` still resolves its path once.
- Error codes: `StyleNotFound`, `FieldNotFound`, `BookmarkNotFound` →
  `not_found`; `SectionNotFound` → `out_of_range`; `BookmarkName`,
  `InvalidProperty` → `invalid_value`.

Tests: `crates/litchi-ops/tests/doc_unified_parts.rs` (11 tests: every
element's reads, writes read back after a reopen, exact no-ops, typed
refusals, the unchanged views, and a corpus sweep that reads every new
kind on every DOC and checks that setting each first paragraph's own style
is byte-exact). One `doc_unified_p0.rs` expectation changed: a missing
hyperlink is now `path_not_found` (pictures keep the gap).

## 4. Deltas from OfficeCLI's DOCX behaviour

| # | OfficeCLI (DOCX) | Litchi (DOC) |
|---|---|---|
| 1 | styles are addressed by `w:styleId` | DOC has no style id; the id is derived from the name (§3) |
| 2 | `/section[N]` and `/body/sectPr[N]` | `/section[N]` (`PlcfSed` order); DOCX in Litchi reads `/body/section[N]` without properties |
| 3 | lengths print in cm | lengths print as other extents do (`EmuConverter.FormatEmu`: cm when exact, else pt); `pos` and level `indent`/`hanging` are integer twips, as record 0319's DOCX keys are |
| 4 | a hyperlink's run formatting is set on its runs | the same: `set` takes `url` and `text`; formatting keys are refused |
| 5 | a new hyperlink defaults to the theme `Hyperlink` style | DOC has no theme: the result is underlined in `#0563C1` over the formatting of the text it is inserted before |
| 6 | `add bookmark` with `text` under a paragraph inserts new text | wraps the first occurrence of `text` in the paragraph; under the body (or the root) it adds a paragraph holding `text` |
| 7 | removing a hyperlink removes its element and runs | removes the whole field (instruction, result, nested fields) |
| 8 | style, numbering, field and TOC writes | read-only here (follow-up 1) |

## 5. Third-party evidence and corpus

- **LibreOffice 26.2.6.3** (`soffice --headless --convert-to docx`) reads every
  written file as intended: the applied paragraph style (`w:pStyle`), the
  run style (`w:rStyle`), the added tab stop, the inserted hyperlink
  (`w:hyperlink` to the URL) after a URL and text change, its removal, the
  added and renamed bookmark, landscape page size and two columns. The
  probe files are written by the ignored test
  `libreoffice_probe_files` (`DOC_WRITE_DIR`).
- **Fidelity** (`results/0289/fidelity.{md,tsv}`): DOC gains a structural
  round trip on every file — a bookmark and a hyperlink added to the first
  body paragraph and both removed — which reads back value-exact on 44 of
  45 files; the other refuses typed (`corrupt_file`, the deliberately
  damaged `litchi-changed/noheadfoot-litchi.doc`, as its text revert
  does). `ole/doc/table-merged-cells.doc`'s revert now finds the cell
  paragraph's `style` key and refuses typed (`not_found`); it found no
  settable key before.
- **CRUD checklist** fixtures: `ole/doc/Lists.doc`, `ole/doc/watermark.doc`,
  `poi/test-data/document/Bug53453Section.doc`,
  `ole/doc/FloatingPictures.doc`, and the generated
  `test-data/unified-ops/0294/marks.doc` (15,360 bytes, SHA-256
  `b2697f64…ce38ca0`, `litchi_doc::Builder`: body tab stops and a text
  box; row in its `PROVENANCE.md`). One recipe adds a hyperlink (it needs
  exactly one of `url` and `anchor`).
- The OfficeCLI script corpus covers DOCX only; its outcomes do not change.

## 6. Measurements

Release builds, `taskset -c 27` on a shared 32-core host (load 7–13).
Evidence: [results/0324/](results/0324/).

**New paths, unified against leaf** (`examples/doc_parts_bench.rs`, on
`test-data/ole/doc/Lists.doc`): each write times one edit against the same
opened owner — leaf: the `doc::Edit` verb and commit; unified: the `File`
verb with path and style-name resolution, admission and the returned
path; both legs publish identical bytes, checked outside the clock. Reads
are warm. 9 rounds after 2 warm-ups, 64 iterations a sample, ABBA,
medians; three trials (`parts-trial-{1,2,3}.json`; binary SHA-256 in
`binaries.sha256`, built by `cargo build --release -p litchi-ops
--features doc --example doc_parts_bench`).

| case | leaf µs | unified µs | unified/leaf |
|---|---:|---:|---:|
| `paragraph_style_set` | 95.5–102.0 | 97.9–103.7 | 1.017–1.025 |
| `run_style_set` | 101.3–106.1 | 105.7–110.0 | 1.037–1.049 |
| `hyperlink_add` | 103.6–111.0 | 108.7–116.5 | 1.047–1.049 |
| `hyperlink_set_url` | 93.2–101.8 | 98.7–109.0 | 1.059–1.070 |
| `bookmark_add` | 84.1–85.3 | 86.1–86.9 | 1.019–1.024 |
| `tab_add` | 94.9–98.1 | 96.4–100.1 | 1.016–1.023 |
| `section_set` | 83.0–84.0 | 83.9–84.9 | 1.011–1.015 |
| `style_get_format` (read) | 0.2 | 0.7 | **3.195–3.397** |
| `paragraph_list_format` (read) | 0.6 | 1.0–1.1 | **1.614–1.649** |
| `hyperlink_get_format` (read) | 2.3 | 2.9 | **1.242–1.296** |
| `section_get_format` (read) | 0.0 | 0.3 | **342–423** |

- Every write is within the 10% single-op budget; the closest is
  `hyperlink_set_url` (+6–7%), whose unified leg also builds the
  hyperlink node (the paragraph's fields and runs) to find the field.
- **Four read misses, not fixed.** The unified point reads cost a fixed
  0.3–0.6 µs: the path is parsed and resolved, a node is built and every
  key of the element's table is projected (a style has 47 keys, a section
  23), where the leaf legs return one typed struct (`section_get_format`'s
  leaf leg copies one struct, 1 ns). The same cost shape as record 0258's
  `run_bold_get` and record 0319's `style_get_format`. Follow-up 7.

**Existing paths, after against before** (`examples/doc_view_regression.rs`,
public `File` API only, built from the same source on a `git archive` of
`db96990ff6` and on the branch; 15 rounds after 3 warm-ups; two runs,
B A A B B A and A B B A A B; `view-regression.txt`):

| scenario | before ns | after ns | after/before (run 1 / run 2) |
|---|---:|---:|---:|
| `/body/p[2]` format (Normal, no list) | 876–956 | 933–968 | 1.046 / 1.071 |
| `/body/p[28]` format (styled, listed) | 923–1,006 | 998–1,031 | 1.058 / 1.092 |
| `/body/p[1]/r[1]` format | 1,388–1,474 | 1,438–1,503 | 1.007 / 1.049 |
| query `/body > p` + every format | 33,709–34,782 | 36,432–37,258 | 1.078 / 1.073 |
| root children | 112–118 | 112–120 | 1.020 / 1.000 |
| body children | 1,516–1,545 | 1,510–1,545 | 1.014 / 0.989 |
| `view text` | 1,907–1,949 | 1,893–1,937 | 0.988 / 0.991 |
| `/body/p[2]` text set | 105,924–111,384 | 107,094–111,959 | 0.996 / 1.017 |
| `/body/p[2]/r[1]` bold set | 95,301–113,476 | 108,678–113,902 | 1.053 / 1.009 |

- A paragraph's `format` now carries seven more keys (`style`,
  `styleId`, `styleName`, `numFmt`, `listStyle`, `start`, `numLevel`) and
  a run's one (`rStyle`); a styled or listed paragraph answers them (new
  output, +6–9%), and a plain one pays the iteration over them (+5–7%,
  about 60 ns). The query over all paragraphs follows (+7–8%). Views and
  writes are unchanged within noise (the bold-set run 1 ratio comes from
  one low base sample).

No claim of faster DOC reads or writes is made.

## 7. Gaps and follow-ups

1. **Style, numbering, field and TOC writes**: `set`/`add`/`remove` of
   `/styles/<id>`, list definitions, instances and levels (an STSH and
   `PlfLst`/`PlfLfo` writer over the revision editor), field and TOC
   `add`/`set`. Paragraph `numId`/`numLevel`/`listStyle` writes belong
   here too.
2. **Revisions** (P2): revision nodes and accept/reject through the
   existing `dispose_revision`.
3. **Form fields**: `FORMTEXT`/`FORMCHECKBOX`/`FORMDROPDOWN` nodes (the
   leaf reads them, `parts::form_fields`) and their writes.
4. **Equations**: EQ fields and MTEF objects as `equation` nodes; writes.
5. **Pictures and OLE objects** as `picture` and `ole` nodes.
6. **Watermark**: the header WordArt shape, read and write.
7. **Read misses** (§6): a projection that formats only the requested
   keys, or a typed fast path for `get(path).format().get(key)`.
8. **Tab stops outside the body** (headers, footers, cells) and
   hyperlink, bookmark and style writes on non-body paragraphs.
9. **Section writes** of `pageNumFmt`, `pageStart`, line numbering,
   direction and the note settings, and section `add`/`remove` (a section
   break insertion).
10. **Text box and shape writes**; shape geometry (`msosptType` from the
    drawing group) and text-box names.

## 8. Gates

Run on the merged tree (trunk `a4a045b48a`); cargo through the
memory-capped shim, `CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`. All exit 0.

- `cargo fmt --all --check`.
- `cargo clippy -p litchi-doc -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings`.
- `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings`, and the same with `--features doc`.
- `cargo test -p litchi-doc --all-features`: 1,487 passed (the
  `source_protection_agrees…` corpus counts gain the generated
  `marks.doc`: 57 → 58 files, 49 → 50 compared).
- `cargo test -p litchi-ops --all-features` in chunks (`--lib`, `--doc`,
  then the 145 integration targets in ten groups of 15, each group's
  executables deleted before the next): 2,446 passed in 148 test binaries.
- `cargo test -p litchi --all-features`: 230 passed.
- `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-doc -p litchi-ops`.
- `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` (42 facade-safe trees, 1 combined tree).
- `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir <wt>/target/glossary` (40 profiles, 2,313 bans, 252 families); `python3 -m unittest tools.test_check_api_glossary`.
- `crates/pyo3-litchi/tests/run_tests.sh`: 282 tests OK.
- `cd tools/perf-baseline && cargo check --locked --all-targets` (no new dependencies).
- `python3 tools/test_no_3rdparty_references.py`; `python3 -m unittest tools.test_officecli_capability_delta` (24 tests); `python3 tools/unified_bench_matrix.py --check`.
- Goldens re-blessed after review, on the merged tree: capabilities and
  registry (`doc.json`; both feature sets: `--all-features` and every
  litchi-ops feature but `encryption`), `results/0289/crud-checklist.{md,tsv}`,
  `results/0289/fidelity.{md,tsv}`, `test-data/unified-ops/0294/marks.doc`
  (`acceptance_fixtures`), `officecli-capability-delta.md` (regenerated,
  `--check` clean).
- The OfficeCLI script corpus (`litchi-cli --test officecli_corpus`) runs
  DOCX scripts only; DOC behaviour is not visible to it, so it was not
  re-recorded.
- Target directory 13 GB after the gates (test executables pruned per group).
