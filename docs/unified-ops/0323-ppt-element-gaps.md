# 0323 — PPT element gaps against the PPTX vocabulary (tier 2)

Date: 2026-10-17. Branch `uo2/ppt-gaps`, cut from trunk
`feat/unified-ops-api` at `db96990ff6`. Owned: `crates/litchi-ppt`,
`crates/litchi-ops/src/format/ppt*`, the PPT rows of the registry,
checklist and fidelity tests. Predecessors: [0260](0260-ppt-p0-surface.md)
(PPT P0 surface), [0321](0321-pptx-element-gaps.md) (the PPTX element
vocabulary this record follows).

The generated delta table listed 18 PPT elements as gaps. This record
closes eleven through the unified `File` API — `comment`, `hyperlink`,
table `row`/`cell`/`column`, `run`, `linebreak`, `theme`, `chart`, `ole`
and `media` (`video`/`audio`/`media`) — and records the rest as deferrals
with follow-ups. PPT element statuses: supported 15 → 26, gap 18 → 7.

## 1. Leaf: `litchi-ppt`

Every write works at the MS-PPT record level. A new root mechanism,
`slide_order::records` (staged by `Transaction::stage_records`), takes the
complete replacement of each persisted record an element edit changes — a
slide's `SlideContainer` and/or the live `DocumentContainer` — computed by
byte-level editors that re-frame only the record path they change; every
other record keeps its bytes. The replacements are published in one
append-only user edit, each replaced record must read back byte-exactly and
the slide list must be unchanged. The change rides the existing
exact/inverse patches and gains a durable operation
(`presentation.element-records.set`: each owner's before-SHA-256 and its
replacement as a blob). Signed sources: an exact no-op keeps every byte, a
changed edit is `Refusal::Signed`. New refusals:
`slide_order::Refusal::{UnsupportedElement, ElementNotFound,
InvalidElementValue}`.

| Element | Read | Write (`package::edit::Edit`, also on `slide_order::Transaction`) |
|---|---|---|
| Comments | `Slide::comments` (existing) | `add_comment` / `set_comment` / `remove_comment` (`comments::{NewComment, CommentChange, DateTime}`): only the slide's `___PPT10` `Comment10Container` records change; a slide without the tag gains `SlideProgTagsContainer` / `___PPT10` / `BinaryTagData` after the records MS-PPT 2.5.1 orders before it. A new comment takes the author's next index, and a document-level `CommentIndex10` author entry (in `DocInfoList`'s `___PPT10` tag) has its seed raised. No clock: an absent date is an all-zero `SYSTEMTIME` |
| Shape hyperlinks | `hyperlink::{Click, Target}`, `Click::resolve`, `Presentation::{shape_click, resolve_click}`: URL, slide (by slide ID, else number) or slide-show jump, with the `ExHyperlink9` screen tip | `set_shape_hyperlink(target, Option<&Click>)`: the shape's `OfficeArtClientData` click `InteractiveInfo` (inserted after `ShapeFlags`/`ExObjRef`/`AnimationInfo`, client data created before `ClientTextbox`, removed when emptied); URL and slide links append an `ExHyperlinkContainer` to `ExObjListContainer` (seed raised; list created after `DocumentAtom` when absent) and, with a tooltip, an `ExHyperlink9Container` to the `DocInfoList` `___PPT9` tag. Superseded definitions are kept, as PowerPoint keeps them |
| Text runs and text hyperlinks | `text_edit::{TextRuns, CharacterRun, TextLink}`, `Slide::shape_text_runs`: the body's `StyleTextPropAtom` character runs (UTF-16 units, covering the final mark) and `InteractiveInfo`/`TextInteractiveInfoAtom` pairs; an outline-backed placeholder reads its slide-list body | `set_run_format(target, start, length, &RunFormat)` (bold, italic, underline, size, sRGB color): the character runs are split at the range ends, the paragraph-run bytes stay exact (a style atom the typed runs would not reproduce byte for byte is refused), a body without a style atom gains one after its text atom. `set_text_hyperlink(target, start, length, Option<&Click>)`: one click pair over exactly the range; an overlapping different range refuses. Outline bodies are edited in the live document |
| Table cells | `TableShape::{cell_shape_id, row_heights, column_widths}` (the grid from the cell anchors) | `set_table_cell_text(target, row, column, text)`: the cell shape's text under the shape-text closure rules; an empty PowerPoint cell (`TextHeaderAtom` only, plus at most one style/special-info/ruler atom) gains its first `TextCharsAtom` with its runs resized from the mark |
| Theme | `theme::{Theme, Fonts}`, `Presentation::theme`: the first master's `RoundTripThemeAtom` theme part (twelve color slots, major/minor fonts, scheme name) | — |

Shared helpers: `record_wire` (byte framing over `text_edit`'s record
walkers) and `binary_tag_edit` (one owner's `___PPT9`/`___PPT10` blob).

## 2. Unified layer: `litchi-ops`

| Path | Read | Write |
|---|---|---|
| `/slide[N]/comment[K]` | `text`, `author`, `initials`, `index`, `date` (`YYYY-MM-DDTHH:MM:SS.fff`), `x`, `y` | `set text/x/y`; `add` (`author` required; `initials` default the word initials; `date`; end placement); `remove` |
| `/slide[N]/shape[M]/hyperlink`, `…/paragraph[K]/run[L]/hyperlink` | `link` (URL, `slide[N]`, `firstslide`/`lastslide`/`nextslide`/`previousslide`/`endshow`), `tooltip`; a run's is the text link holding its start | `add` (`link` alias `url`, `tooltip`) under a shape or run; `set` (`link=none` removes, tooltip-only keeps the target); `remove`. OfficeCLI's scheme allow-list (`javascript:` … refused); other `ppaction://` actions refuse |
| `/slide[N]/table[M]/tr[R]`, `…/tr[R]/tc[C]`, `…/col[C]` | row `height`, column `width`, cell text | cell `text`; row/column `add` and `remove` refuse (`UnsupportedCapability`, operation gaps) |
| `…/paragraph[K]/run[L]`, `…/paragraph[K]/linebreak[B]` | run text, `font.latin`, `size`, `bold`, `italic`, `underline` (`single`/`none`), `color` (sRGB); line breaks (`\v`) | run `bold`, `italic`, `underline`, `size`, `color`, `text` (same length, or a single-run body) |
| `/theme` | the PPTX theme keys (`dk1` … `followedhyperlink`, `headingFont`, `bodyFont`, `.ea`/`.cs`, `name`) | — (gap) |
| `/slide[N]/chart[M]` | frame `id`, `name`, `x`, `y`, `width`, `height`; `chartType` (PPTX names from the MS-OGRAPH group family), `title` | — (gap) |
| `/slide[N]/ole[M]`, `video[M]`, `audio[M]`, `media[M]` | frame keys; OLE `progId`, `linked`; media kind from the external-media object (movie → `video`) | — (gaps) |

The new elements are reached by path and by a query that names them; views
and `*` walks list the same children as before (the 0321 policy), so
existing views and query results keep their shape. The PPT write
dispatch, the settable/addable/removable lists and the capability property
tables moved into `format::ppt` (`SETTABLE`, `ADDABLE`, `REMOVABLE`,
`capability_props`, `WRITE_GAPS`); `capabilities.rs` keeps one-line arms.

## 3. Deferrals and follow-ups

| Gap | Follow-up |
|---|---|
| Table row/column insertion and removal (cell shapes and border lines re-laid out) | 0323-F1 |
| Nested-group tables, merged-away cells, rich cell text | 0323-F2 |
| LibreOffice resolves PPT slide-jump hyperlinks only through the `_PID_HLINKS` DocumentSummaryInformation property (URL links render); the leaf does not write it | 0323-F3 |
| Run and line-break insertion and removal (text positions move) | 0323-F4 |
| `animation` nodes (the timing tree is read by the leaf, not projected) | 0323-F5 |
| `equation` (Equation Editor OLE objects) | 0323-F6 |
| Chart series and axes; chart, series and axis writes | 0323-F7 |
| OLE writes | 0323-F8 |
| Media writes (movie/sound objects and frames) | 0323-F9 |
| Theme writes (re-encoding the `RoundTripThemeAtom` packages and the legacy color schemes) | 0323-F10 |
| Comment author list is not extended for a new author (PowerPoint reads comments without it) | 0323-F11 |
| `zoom`, `model3d`, `modernComment`, `reply` | wontfix: PPTX-only OOXML objects |
| Read misses (§5) | 0323-F12 |

## 4. Evidence

| Check | Result |
|---|---|
| Leaf tests | `element_comments` (5: tag creation, existing comments, seed raise, no-op, durable replay/revert, strict dates), `element_tables` (2: length change keeps the grid, empty PowerPoint cells gain a text atom), `element_hyperlinks` (2: URL with tooltip, slide, jump, replace, remove, refusals), `element_runs` (5: PowerPoint runs, split/round trip/no-op, outline body format, text links, WithLinks outline links read), `element_theme` (2); every write checks the exact inverse restores the source bytes |
| Ops tests | `tests/ppt_elements.rs` (12): every element on the eager and positional owners after a byte round trip, the source left unchanged, exact no-ops, typed refusals |
| LibreOffice 25 (`soffice --convert-to odp`) | a comment added (author, text), a URL shape hyperlink and a table-cell text set read back in `content.xml`; a slide-jump link does not (F3) |
| CRUD checklist | re-blessed; fixtures `unified-ops/0294/elements.ppt` (generated, provenance in that directory's `PROVENANCE.md`), POI `headers_footers_2007.ppt`, `testPPT_oleWorkbook.ppt`, `sound.ppt`; recipes for comment and hyperlink `add`; every claimed PPT cell passes; `video` is unreached (no movie fixture) |
| Fidelity | re-blessed: the new deck is `exact`; no other row changed |
| Goldens | `capabilities/ppt.json`, `registry/ppt.json` re-blessed with all features and with every feature but `encryption` (identical PPT bytes; no other format changed) |
| Delta table | PPT elements supported 15 → 26, gap 18 → 7; property reads supported 41 → 94; property writes supported 0 → 6, partial 14 → 23 |

## 5. Measurements

Probe `tools/perf-baseline/examples/ppt_elements_probe.rs` (release, no
LTO, built from `250366151e` plus the probe, binary SHA-256
`1113842cb5f379d67f8321dbc140a67aba4cc6ae9d3d516cf2f64003636ee4dd`, not
committed), `taskset -c 7`, warm owners, 9 ABBA rounds (18 samples a leg),
two runs on a shared host (load 11–12):
[`ppt-elements.json`](results/0323/ppt-elements.json),
[`ppt-elements-repeat.json`](results/0323/ppt-elements-repeat.json). Every
write route publishes the same bytes on both legs.

| Route | Slides | Native median (run 1 / 2) | Unified median | Ratio (run 1 / 2) |
|---|---:|---:|---:|---|
| `ppt_comment_set_v1` | 1 | 105.2 / 106.3 µs | 110.3 / 110.4 µs | 1.048 / 1.040 |
| `ppt_comment_set_v1` | 50 | 373.8 / 375.7 µs | 375.4 / 382.9 µs | 1.016 / 1.021 |
| `ppt_comment_add_v1` | 1 | 107.9 / 107.8 µs | 111.1 / 110.5 µs | 1.023 / 1.025 |
| `ppt_comment_add_v1` | 50 | 449.9 / 451.0 µs | 462.4 / 461.0 µs | 1.029 / 1.017 |
| `ppt_cell_text_set_v1` | 1 | 106.4 / 105.9 µs | 111.2 / 110.6 µs | 1.045 / 1.042 |
| `ppt_cell_text_set_v1` | 50 | 354.5 / 355.9 µs | 363.4 / 361.7 µs | 1.038 / 1.034 |
| `ppt_hyperlink_add_v1` | 1 | 145.1 / 144.6 µs | 150.8 / 147.7 µs | 1.040 / 1.026 |
| `ppt_hyperlink_add_v1` | 50 | 472.1 / 468.0 µs | 479.1 / 476.6 µs | 1.016 / 1.022 |
| `ppt_run_format_v1` | 1 | 107.8 / 107.6 µs | 113.2 / 112.9 µs | 1.054 / 1.049 |
| `ppt_run_format_v1` | 50 | 362.8 / 355.8 µs | 368.4 / 369.2 µs | 1.009 / 1.052 |
| `ppt_comment_get_v1` | 1, 50 | 1.3 µs | 1.9 µs | 1.43–1.45 |
| `ppt_run_get_v1` | 1, 50 | 0.7 µs | 2.1–2.2 µs | 3.06–3.11 |

Every write is within the single-op budget (≤ 10%). **Misses:** the two
reads. Both are microsecond-scale; the unified leg adds path parsing, the
resolution walk and the node's step array, and the run read builds the
paragraph's pieces (the native leg reads the first character run only).
The added cost is 0.6 µs and 1.5 µs, flat in deck size (proportional to the
path, as the read budget asks), but the ratio misses. Follow-up 0323-F12:
cache a slide's comments and a shape's runs per node handle, and pair the
run leg with the same piece construction.

## 6. Verification (after merging trunk `fde847fe90`)

| Gate | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-ppt -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features ppt] --all-targets -- -D warnings` | 0, 0 |
| `cargo test -p litchi-ppt --all-features` (1,499 tests) | 0 |
| `cargo test -p litchi-ops --all-features` in `--lib` (922), `--tests` (2,236), `--doc` (210) chunks | 0 |
| `cargo test -p litchi --all-features` (230) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-ppt -p litchi-ops` | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0, 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (40 profiles, 2,313 bans, 252 families) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (282 tests) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| goldens (`LITCHI_BLESS=1 … --test capabilities --test registry`, both feature sets), `python3 tools/officecli_capability_delta.py --officecli …/3rdparty/OfficeCLI`, `PYTHONPATH=. python3 tools/test_officecli_capability_delta.py` | 0 |
| `cargo test -p litchi-cli --all-features --test officecli_corpus` (PPT is outside the corpus; unchanged, not re-blessed) | 0 |

No code, test or tool references `3rdparty/`; every fixture is under
`test-data/`.

## 7. Not claimed

- Rendering in PowerPoint; LibreOffice was checked for the cases in §4.
- Byte parity with PowerPoint's own serialization of new comments, links
  and style atoms; the written records follow the MS-PPT layouts above.
- Frozen-harness qualification; the probe is a local paired measurement.
