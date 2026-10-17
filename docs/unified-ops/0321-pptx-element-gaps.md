# 0321 — PPTX element gaps for OfficeCLI tier-1 parity

Date: 2026-10-16. Branch `uo2/pptx-p1`, cut from trunk
`feat/unified-ops-api` at `c3d0a2a858`.

The generated delta table (`officecli-capability-delta.md`) listed eight
PPTX elements as gaps. This record closes five of them through the unified
`File` API — `theme`, `column` (`col`), `hyperlink`, `media`
(`video`/`audio`/`media`) and `equation` — adds writes to slide masters
and layouts, and records `animation`, `zoom` and `model3d` as P2
deferrals. PPTX element statuses in the delta table: supported 25 → 30,
gap 8 → 3.

## 1. What landed

### Leaf: `litchi-pptx`

| Area | Read | Write |
|---|---|---|
| Theme | `Package::theme`, `SourceBackedPresentation::theme`, `shape::theme::presentation_theme` (the presentation part's own theme, else the first master's: OfficeCLI's `GetThemePart`) | `opened::Edit::{theme, set_theme_colors, set_theme_fonts}` over the existing `put_colors`/`put_fonts` scheme splices; a palette naming the same colors in any slot order, or an equal font set, is an exact no-op |
| Masters and layouts | (existing catalog reads) | `Edit::{set_master_name, set_layout_name}` (`p:cSld/@name`, the slide-name lexical editor generalized to the `sldMaster`/`sldLayout` roots); `Edit::{set_master_shape_text, set_layout_shape_text}` (the slide shape-text rewriter on master and layout parts) |
| Table columns | `table::Table::grid_columns` (`a:gridCol/@w`, in the owner's namespace scope) | `Edit::{set_table_column_width, insert_table_column, remove_table_column}` (`table::columns`): only the `w` attribute, the inserted or removed `a:gridCol` and its cell in every row change, and the frame's `p:xfrm/a:ext/@cx` follows the sum of the widths (OfficeCLI's frame sync); tables with cells merged across columns refuse insertion and removal, the last column is kept |
| Hyperlinks | `hyperlinks::{Click, Link, Target}`: `Click::of_shape` (a shape's `p:cNvPr/a:hlinkClick`), `Click::from_text` (a run's); `Slide::hyperlink`, `SourceSlide::hyperlink`, `source::Slide::hyperlink` resolve a click through the slide's relationships to `Target::Url`, `Target::Slide(Position)` or a relationship-free `Target::Action` | `Edit::{set_shape_hyperlink, set_run_hyperlink}`: a shape-wide link goes on `p:cNvPr` and every text run (OfficeCLI's `ApplyShapeHyperlink`), in `CT_TextCharacterProperties` order; external targets reuse an existing hyperlink relationship to the same URL, slide jumps the slide relationship; an equal resolved link is an exact no-op; relationships a removal leaves unused are kept |
| Media | `media_parts::{Summary, summary}`: a picture's kind (`a:videoFile`/`a:quickTimeFile`, `a:audioFile`/`a:wavAudioFile`/`a:audioCd`) and PowerPoint 2010 trim, from the picture's own fragment | `Edit::add_media` with `MediaFile`s for the media and the poster: the media part, its `video`/`audio` and `media` relationships, the poster and its image relationship (`media_parts::store_with`, which lifts `store`'s one-batch-per-slide rule for appends) |
| Equations | `shape::math::{Equation, equations}`; `Slide::equations`, `SourceSlide::equations`, `source::Slide::equations` read the slide's own bytes (the shape scene reads the `mc:Fallback` picture), each `a14:m` with its shape's `cNvPr` ID, a LaTeX-like `formula` and its text | — |

### Unified layer: `litchi-ops`

| Path | Read | Write |
|---|---|---|
| `/theme` | `dk1`…`followedhyperlink` (`#RRGGBB`; a system color reads as its last RGB), `headingFont`, `bodyFont`, `.ea`/`.cs` variants, `name` (the color scheme's) | `set` of every key, with OfficeCLI's aliases (`dark1`, `hlink`, `majorFont`, `minor`, …); `none`/`default` clear a typeface |
| `/slidemaster[N]`, `/slidemaster[N]/slidelayout[M]`, `/slidelayout[M]` | (unchanged) | `set name`; `set text` on a direct shape or placeholder of a master or layout |
| `/slide[N]/table[M]/col[C]` | `width` | `add` (`width`, default the average; seed `text`; start, end or position placement), `set width`, `remove` |
| `/slide[N]/shape[M]/hyperlink`, `…/paragraph[K]/run[L]/hyperlink` | `link` (URL, `slide[N]`, `firstslide`/`lastslide`/`nextslide`/`previousslide`/`endshow`), `tooltip`; the shape's is its `p:cNvPr` link, else its first run's | `add` (`link` alias `url`, `tooltip`) under a shape or run; `set` (`link=none` removes; a tooltip-only set keeps the target); `remove`. `javascript:`, `vbscript:`, `data:`, `file:` and `jar:` targets are refused (OfficeCLI's allow-list) |
| `/slide[N]/video[M]`, `audio[M]`, `media[M]` | the frame keys, `trimStart`, `trimEnd` (milliseconds) | `add` (`src` alias `path`, `poster` required, `name`, `x`, `y`, `width`, `height`; `media` takes its kind from the file); `remove` |
| `/slide[N]/shape[M]/equation[K]`, `/slide[N]/equation[K]` | `formula`; text | — |

Media pictures remain pictures: `picture[K]` still counts them and the
slide's view lists them as pictures. The new elements are reached by path
and by a query that names them; the node view and `*` queries list the
same children as before, so existing views and query results keep their
shape. `capabilities()` and `schema()` report every new element, its keys
and its verbs; the PPTX lists of settable, addable and removable elements
moved from `capabilities.rs` into `format::pptx` (`SETTABLE`,
`SHAPE_ADDABLE`, `REMOVABLE`, `capability_props`).

Source owners are promoted to an eager package for every new write, as for
charts and SmartArt (the source-backed edit stages slide text and topology
only).

## 2. OfficeCLI deltas

- `media add` requires a `poster`; OfficeCLI generates one. The media
  payload is stored inert (never decoded); its type comes from the file
  extension or the data URI's media type.
- `formula` is LaTeX-like: fractions, scripts, radicals, n-ary operators
  and delimiters are reconstructed; other OMML structures read as their
  text. OfficeCLI's own reconstruction may spell some differently.
- Hyperlinks: `add` also accepts a run parent (OfficeCLI adds shape links
  only; run links are its `set run link=`).
- Theme: color sets accept opaque RGB only (as OfficeCLI's six-digit hex
  rule).

## 3. Gaps and follow-ups

| Gap | Follow-up |
|---|---|
| `shape.link`/`tooltip` and `run.link`/`tooltip` as shape and run properties (OfficeCLI's canonical write path) | 0321-F1: route them to the hyperlink writes; reading them changes the shape and run tables, so it waits for a coordinated re-record |
| Picture and group click actions | 0321-F1 |
| Equation `set`, `add`, `remove` | 0321-F2: an OMML writer in `litchi-drawingml` (OfficeCLI's `FormulaParser` input) |
| Media `set`, `volume`, `autoPlay`, `loop` (timing tree) and `src` replacement | 0321-F3: media timing nodes in `animations` |
| Master/layout `background`, `direction`, and shape properties other than `text` | 0321-F4 |
| `table.colWidths`, `presentation.defaultFont`, `picture.mediaType` | 0321-F5: added to their tables in one re-record |
| `animation` (P2) | M1-extended: the leaf `animations` codec reads the timing tree; the adapter needs per-shape targets |
| `zoom` (P2) | M1-extended: `shape::zoom::read` already reads zoom owners from slide XML; the adapter needs the slide's raw bytes as for equations |
| `model3d` (P2) | M1-extended: `model3d` codec reads the frames; same adapter step |

## 4. Evidence

| Check | Result |
|---|---|
| Unit, `litchi-pptx` | theme colors and fonts round trip, semantic no-op; master and layout renames; column width, insertion (with seed text, escaped) and removal keep the frame width in step, merged tables and the last column refused; click-slot scanner (schema order, `rtl`/`extLst` insertion point); link resolution of relationships and actions; shape and run links round trip (URL with tooltip, slide jump, action, removal, no-op after reopen, out-of-range run); two media pictures on one slide, kind mismatch refused; the equation of `import-mce.pptx` read from the choice branch (`𝑎=\frac{𝑏}{𝑐}`) |
| Integration, `crates/litchi-ops/tests/pptx_elements.rs` | every element on source and eager owners after a byte round trip: get, query, set, add, remove, exact no-ops, the input left unchanged, typed refusals (unknown key, non-RGB theme color, slide jump past the end, `javascript:` link, missing `link`/`poster`, kind mismatch, out-of-range column, negative width, a master shape's non-text property, equation `set`); `*` queries list no new kinds |
| Capability and registry goldens | `tests/golden/{capabilities,registry}/pptx.json` re-blessed with all features and with every feature but `encryption`; no other format changed; `tests/capabilities.rs` PPTX rows updated |
| Delta table | regenerated: PPTX elements supported 25 → 30, gap 8 → 3 (`animation`, `model3d`, `zoom`); property reads supported 197 → 226 |
| CRUD checklist (`results/0289/crud-checklist.*`) | re-blessed: recipes for `video`, `audio`, `media` add; fixtures `import-mce.pptx` (LibreOffice, MPL-2.0) and POI `EmbeddedAudio.pptx`/`EmbeddedVideo.pptx` (Apache-2.0, already tracked) reach equations and media; no PPTX cell unreached or failing |
| Fidelity (`results/0289/fidelity.*`) | re-blessed: every PPTX file stays `exact` on open → empty edit → save; the set-then-revert round trip now reaches a layout `name` on 74 decks (was refused or absent) and reverts byte-exactly; one (`poi/.../bug62513.pptx`) is one-way: the set back is refused because the promoted package cannot preserve the owned source (typed, as `49609.xlsx`) |
| Python | `test_write_api.PptxVerbTests.test_record_0321_elements`; the stubs type props as dictionaries (`_Props`), so no stub change; `run_tests.sh`: 280 tests OK |

### OfficeCLI corpus

Re-recorded (`LITCHI_BLESS=1 cargo test -p litchi-cli --all-features
--test officecli_corpus`, `python3 tools/officecli_corpus.py`, `--check`
clean). Passed calls are unchanged (1286 PPTX); what changed is how the
remaining failures are classified:

| Family | Gap | Delta | Cascade | Warned |
|---|---|---|---|---|
| ppt | 143 → 141 | 0 → 2 | 1 → 1 | 37 → 37 |
| ppt/tables | 69 → 65 | 0 → 0 | 286 → 290 | 7 → 7 |
| pptx total | 416 → 410 | 144 → 146 | 397 → 401 | 716 → 716 |

`add column` (4 calls, `tables-rows-cols.sh`) is no longer an element gap:
it now fails as a cascade because the script's earlier `add table` is
refused (PPTX table insertion is not a unified verb yet). `add video` (2
calls, `video.sh`) is a delta: the harness does not ship the script's media
file (`file_not_found`); the call's `volume`, `autoplay`, `loop` and
`trimStart` are skipped with `--unsupported warn` (timing-tree gap F3).

## 5. Measurements

Harness `tools/perf-baseline/unified-pptx-elements` (README there):
release build (`CARGO_NET_OFFLINE=true cargo build --release
--manifest-path tools/perf-baseline/unified-pptx-elements/Cargo.toml`,
source tree at `c3d0a2a858` plus this change, binary SHA-256
`1dc58f885385c5194006994c0a184cef398f375a85a45e671d1579ec20df4186`, not
committed), pinned with `taskset -c 7`, 9 ABBA rounds (18 samples per
leg), eager owners, two runs on a shared machine (load average 5–6)
([`pptx-elements.json`](results/0321/pptx-elements.json),
[`pptx-elements-repeat.json`](results/0321/pptx-elements-repeat.json), raw
samples included). Every write route publishes the same bytes as its
native leg.

| Route | Slides | Native median | Unified median | Ratio (run 1 / run 2) |
|---|---:|---:|---:|---|
| `pptx_theme_set_v1` | 1 | 0.640 / 0.641 ms | 0.645 / 0.643 ms | 1.008 / 1.003 |
| `pptx_theme_set_v1` | 50 | 2.117 / 2.135 ms | 2.132 / 2.146 ms | 1.007 / 1.005 |
| `pptx_master_name_v1` | 1 | 0.469 / 0.471 ms | 0.474 / 0.473 ms | 1.010 / 1.004 |
| `pptx_master_name_v1` | 50 | 1.846 / 1.857 ms | 1.849 / 1.854 ms | 1.002 / 0.998 |
| `pptx_column_width_v1` | 1 | 0.567 / 0.568 ms | 0.593 / 0.598 ms | 1.045 / 1.052 |
| `pptx_column_width_v1` | 50 | 1.954 / 1.969 ms | 1.989 / 1.999 ms | 1.018 / 1.016 |
| `pptx_column_add_v1` | 1 | 0.586 / 0.595 ms | 0.641 / 0.651 ms | 1.094 / 1.094 |
| `pptx_column_add_v1` | 50 | 1.983 / 2.002 ms | 2.059 / 2.056 ms | 1.038 / 1.027 |
| `pptx_hyperlink_set_v1` | 1 | 0.549 / 0.549 ms | 0.556 / 0.555 ms | 1.011 / 1.010 |
| `pptx_hyperlink_set_v1` | 50 | 1.992 / 2.007 ms | 2.006 / 2.008 ms | 1.007 / 1.000 |
| `pptx_hyperlink_get_v1` | 1 | 30.9 / 31.4 µs | 5.5 / 5.4 µs | 0.176 / 0.171 |
| `pptx_hyperlink_get_v1` | 50 | 86.2 / 88.2 µs | 6.0 / 5.9 µs | 0.070 / 0.067 |
| `pptx_equation_get_v1` | 1 | 47.8 / 46.8 µs | 62.4 / 60.6 µs | 1.304 / 1.295 |
| `pptx_theme_get_v1` | 1 | 53.8 / 54.6 µs | 53.3 / 54.3 µs | 0.991 / 0.996 |
| `pptx_theme_get_v1` | 50 | 63.3 / 62.9 µs | 63.1 / 63.1 µs | 0.997 / 1.002 |

Every write is within the single-op budget (≤ 10%). `pptx_column_add_v1`
on a one-slide deck is the closest (1.094): the unified add reads the grid
once more to place the column and bound the seed text, a fixed ~55 µs that
the 50-slide row dilutes. The hyperlink read is faster through the unified
handle because the eager handle keeps the slide's scene and text body,
which the native leg parses on every read; it is not a like-for-like
saving. **Miss:** `pptx_equation_get_v1` is 1.30×: the native leg calls
`Slide::equations` alone, while the unified path also resolves
`/slide[1]/shape[1]` (the slide's scene) and matches the shape's ID; the
read is ~14 µs dearer in absolute terms. Follow-up 0321-F6: cache a
slide's equations per handle (as charts are) so repeated equation reads
skip the slide scan, and pair the native leg with the same shape lookup.
Source-backed owners are not measured: every new write promotes a source
owner to an eager package first, as chart and SmartArt writes do, which
the native API has no counterpart for.

## 6. Not claimed

- Rendering of the new writes in PowerPoint or LibreOffice was not run.
- OfficeCLI byte parity of the written XML; the written markup follows the
  OOXML schema order and OfficeCLI's placement rules, not its serializer.
