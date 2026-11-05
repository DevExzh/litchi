# 0270 — Typed chart CRUD through the unified API

Date: 2026-10-13. Status: implements record 0248 §6 Phase E step 19 (the
chart half; SmartArt is not part of this record). Branch `uo2/s19-charts`,
cut from trunk `feat/unified-ops-api` at `8fd16dd11e`.

PPTX slide charts and XLSX worksheet charts can now be added, read, edited
and removed through `File::{add, get, set, remove}` (and therefore
`edit`, `batch`, Python and the CLI), with their series and axes as nodes.
DOCX charts remain a gap (§6).

## 1. Inventory before this change

| Owner | Could author | Could edit | Could read |
|---|---|---|---|
| `litchi-drawingml::chart` | a complete typed `Chart` model and `writer::write` (every classic family, 3D, axes, labels) | `edit::Edit`: title, language, style, blanks, series title, data-label flags, axis range/position (byte-range, source-preserving) | `reader::read` to the typed model |
| `litchi-pptx` | `opened::Edit::add_chart` from a small `chart::Chart` model with its own hand-written XML, no embedded workbook | nothing | `chart_info` (type, title, legend flag); frames counted in the scene |
| `litchi-pptx::presentation::embedded::excel` | `Workbook::from_chart`: a deterministic XLSX for chart data, never wired to a chart | — | — |
| `litchi-xlsx` | chartsheets only; no worksheet chart authoring on `Package` | drawing transfer when a sheet is copied | drawing inventory (`Drawing::charts`), no chart-part read on the owners |
| `litchi-spreadsheet-drawing` | `Chart::{bar,line,pie,area,scatter}_chart` from range formulas (XLSB writer) | — | — |
| `litchi-docx` | `chart::Graph` store of chart parts and embedded workbooks; no inline drawing authoring | `chart::Transaction` over the graph | `Package::chart_graph` |
| `litchi-ops` | nothing (pptx `add chart` → `unsupported_type`; xlsx chart elements were gaps L-XLSX-R7) | nothing | pptx chart frames with frame properties only; `chartType` was a property gap |

## 2. What landed

### Leaf: `litchi-drawingml`

- `chart::authoring` (new): `Spec` describes a chart the way a caller
  thinks of it — a `Variant` (`Kind` column, bar, line, pie, doughnut,
  area, scatter, bubble, radar, stock or combo; `Grouping` standard,
  stacked or percent-stacked; 3D), title, categories, named `Series`,
  legend position, category and value axis titles, an `Appearance`
  (colors incl. ARGB, data labels and their position, vary colors, gap
  width, overlap, hole size, first slice angle, scatter and radar style,
  smoothing, explosion, legend overlay, auto-title deletion, value-axis
  min/max/major unit, major and minor gridlines with color/width/dash,
  markers, line width, 3D view) and a `Binding`. `Spec::to_xml` writes a
  deterministic chart-space part through the existing writer;
  `Binding::Literal` writes `c:numLit`/`c:strLit`, `Binding::Sheet`
  writes references with caches into the layout `Spec::sheet_layout`
  describes, at any origin. `Variant::of(&Chart)` classifies a read chart.
  Invalid combinations (radar 3D, stacked pie, stock without three or four
  series, labels at a position the chart group's schema omits, ...) are
  `Error::Invalid`; series, points, text and worksheet bounds are
  `Error::Limit`.
- `chart::edit::Edit` gains `set_legend`, `set_axis_title`,
  `set_series_values` and `set_categories`. Like the existing verbs they
  rewrite only the selected element's bytes; a reference keeps its formula
  and gets a new cache; a single-column reference is resized when the
  point count changes, any other reference refuses a resize.

### Leaf: `litchi-pptx`

- `opened::Edit::add_chart(slide, &Spec, name, bounds) -> shape id`
  (breaking: the old `&chart::Chart` argument is gone; ADR 0001) stages the
  chart part, a deterministic embedded workbook
  (`embedded::excel::Workbook::from_spec`, `Sheet1` in the spec's layout),
  the chart → package and slide → chart relationships, and the graphic
  frame, in one candidate that becomes visible only when every step
  succeeded.
- `opened::Edit::{chart, set_chart}` read and replace the chart part (and,
  when given, the embedded workbook bytes) of a chart frame selected by
  its depth-first shape position; identical bytes are an exact no-op.
- `slide::Slide::chart_xml`, `presentation::SourceSlide::chart_xml` and
  `source::Slide::chart_xml` read one chart part by the frame's payload
  relationship, so a chart node decodes only its own part.
- Removal is the existing `remove_shape`: the eager owner drops the chart
  part and its workbook with the frame; the source-backed edit keeps unused
  parts for exact source replay (its documented policy).

### Leaf: `litchi-xlsx`

- `chart::edit` (new scoped family, registered in the glossary gate):
  `Package::edit_charts() -> Edit`; `Edit::add(sheet, &Anchor, name, xml)`
  creates the drawing part and the worksheet's `drawing` element when the
  sheet has none, the chart part and relationship, and a two-cell anchor;
  `Edit::{chart, set, remove}`; `commit() -> Commit` with
  `package/into_package/patch/diagnostics/into_parts`; `Patch` is the exact
  part-level before/after set (`is_noop`, `inverse`), replayed by
  `Package::apply_charts` with `PatchConflict` on any other snapshot.
  Charts are numbered per sheet in drawing order (ordinary `c:chart`
  frames). A changed edit of a signed workbook is `Error::Signed`.
- `Package::{chart_count, chart_xml}` and
  `SourceWorksheet::{chart_count, chart_xml}` read without materializing
  the workbook; the source owner reads only the worksheet relationships,
  the drawing and the selected chart part.

### Unified layer: `litchi-ops`

- `format::chart` (shared by PPTX and XLSX): OfficeCLI's `chartType`
  vocabulary (`column`, `stackedBar`, `percentStackedColumn3d`, `donut`,
  `xy`, `combo`, ...), `legend` values (`b`, `topRight`, `none`, `true`),
  the `data` grammar (`Name:1,2,3;Other:4,5,6`, and `Q1:40,Q2:55`),
  `series1`..`series20`, 26 formatting keys, the `Summary` read projection
  and `ChartEdit` (one source-preserving chart edit per `set`).
- Registry: `chart` is a real element on both formats (add, set, get,
  query, remove), `series` (`name`, `values` settable; `categories`
  read) and `axis` (`role` read, `title` settable; `axis[@role=value]` or
  `axis[N]`) replace the gaps. Chart reads: `chartType`, `title`,
  `legend`, `categories`, `catTitle` (PPTX adds the frame properties,
  XLSX `seriesCount`). Add-only: `data`, `axisTitle` (`vtitle`), the
  formatting keys and `series{N}`; XLSX also `dataRange`, `name` and
  `x/y/width/height` in cells (OfficeCLI's units).
- PPTX add/set promote a source owner to an eager package (as picture add
  does); a data or series-name `set` regenerates the embedded workbook when
  the chart refers to the authored layout, and a data `set` of any other
  embedded workbook is refused (`unsupported_capability`), so the workbook
  never silently diverges from the cached values.
- XLSX `dataRange=Sheet1!A1:E13` reads the host cells (names in the first
  row, categories in the first column, numbers below; formula caches are
  used) and binds the chart to them with cached values; inline `data` is
  written as literals, so no workbook part is needed.

## 3. Evidence

| Check | Result |
|---|---|
| Unit tests, `litchi-drawingml::chart::authoring` | every variant × grouping × 3D × binding round-trips through the reader, deterministic bytes, sheet layout and references, appearance, combo, invalid specs; edit verbs on references and literals |
| Unit tests, `litchi-pptx::opened` | deterministic `add_chart` (equal package bytes twice), embedded workbook present, `chart`/`set_chart` no-op and replace, invalid XML refused, removal drops chart and workbook |
| Unit tests, `litchi-xlsx::chart::edit` | add/set/remove round trip, determinism, replay, inverse, `PatchConflict`, no-op, invalid input |
| Integration, `crates/litchi-ops/tests/charts.rs` | PPTX and XLSX add/get/set/remove for source and eager owners after a byte round trip; series and axis paths; clearing a title; no-op `set` (`patch.is_noop()`, equal bytes); refusals (missing data, bad value, unknown and unsupported types, unknown keys, set-only keys) leave the input unchanged; batch add+set, replay and inverse; signed PPTX and XLSX refuse (`document_protected`) with the input unchanged, and the unsigned PPTX accepts; encrypted PPTX refuses plaintext publication; capability flags |
| Goldens | `tests/golden/{capabilities,registry}/{pptx,xlsx,xlsx-encryption}.json` re-blessed: chart/series/axis become readable, settable, addable and removable; no other element changed |
| LibreOffice 26.2.6.3 headless | 20 variants (`column` … `percentStackedLine`, `combo`, 3D, stock) authored through the CLI in one PPTX and one XLSX, converted to PDF; every chart title is rendered ([`results/0270/lo-check.json`](results/0270/lo-check.json), script [`results/0270/lo_check.sh`](results/0270/lo_check.sh)) |

## 4. OfficeCLI example corpus

`LITCHI_BLESS=1 cargo test -p litchi-cli --test officecli_corpus`, then
`python3 tools/officecli_corpus.py`. Only chart rows of `outcomes.tsv`
changed (789 rows; no other row). Passed calls went from 2,003 to 2,196;
no call regressed; the bug count stays 0.

| Chart `add` lines | Before | Now |
|---|---:|---:|
| PPTX | 0 / 447 (`unsupported_type`) | 155 / 447 |
| XLSX | 0 / 322 (`unsupported_capability`, L-XLSX-R7) | 35 / 322 |
| DOCX | 0 / 14 | 0 / 14 (gap, §6) |

The remaining chart lines fail on keys this record does not author yet
(each line stops at its first unsupported key): `preset` (37), font specs
`labelfont`/`legendFont`/`title.font`/`axisfont` (47), fills and effects
(`gradient(s)`, `transparency`, `chartareafill`, `plotFill`,
`seriesoutline`, `series.shadow`; ~60), `secondaryAxis` (21),
`referenceLine` (14), `labelrotation` (8), `trendline` (6), and the
chart types `waterfall` (47 lines, OfficeCLI simulates it with stacked
bars), `histogram`, `boxWhisker`, `funnel`, `treemap`, `pareto`,
`sunburst` (51 lines; `cx:chart` extended charts). Many XLSX chart lines
also follow cell writes that use the undocumented `text` alias (G7, owned
by the corpus-parity stream); those lines pass once it lands. New delta
D9 records that a chart-format key the chart group's schema omits
(`labelPos` on an area chart) refuses the whole `add` instead of being
dropped with a warning.

## 5. Measurements

Harness `tools/perf-baseline/unified-chart-add` (README there): release
build (`CARGO_NET_OFFLINE=true cargo build --release --manifest-path
tools/perf-baseline/unified-chart-add/Cargo.toml`, binary SHA-256
`7ed587bcf1d8e771cd0246e6aceadb5fe2b9cfc14839446528b0a024ef9cb161`, not
committed), pinned with `taskset -c 7`, 9 ABBA rounds (18 samples per leg),
eager owners holding `N` charts, one chart add per sample. Both legs publish
the same package (XLSX exact bytes; PPTX equal decoded members — the member
order differs between the two owners' writers). The machine was shared with
other agents (load average 9–17 on 32 cores), so two runs are recorded
([`chart-add.json`](results/0270/chart-add.json),
[`chart-add-repeat.json`](results/0270/chart-add-repeat.json)), with
raw samples.

| Route | N | Native median | Unified median | Ratio (run 1 / run 2) | Native / unified minimum |
|---|---:|---:|---:|---|---|
| `pptx_chart_add_v1` | 0 | 1.10 / 0.97 ms | 0.98 / 0.99 ms | 0.89 / 1.02 | 0.931 / 0.937 ms |
| `pptx_chart_add_v1` | 10 | 4.56 / 4.59 ms | 4.61 / 1.63 ms | 1.01 / 0.36 | 1.571 / 1.588 ms |
| `pptx_chart_add_v1` | 100 | 12.52 / 12.74 ms | 12.52 / 12.81 ms | 1.00 / 1.01 | 12.46 / 12.48 ms |
| `xlsx_chart_add_v1` | 0 | 47.1 / 47.0 µs | 49.3 / 49.6 µs | 1.05 / 1.06 | 45.9 / 48.4 µs |
| `xlsx_chart_add_v1` | 10 | 119.2 / 120.0 µs | 120.1 / 121.2 µs | 1.01 / 1.01 | 116.1 / 118.0 µs |
| `xlsx_chart_add_v1` | 100 | 776.7 / 767.2 µs | 770.2 / 768.0 µs | 0.99 / 1.00 | 751.4 / 748.2 µs |

The unified add is within the single-op write budget (≤ 10%) on every
case; the XLSX overhead at `N = 0` is about 2.5 µs (property admission,
`data` parsing and the spec build), and the PPTX/10 medians are bimodal
between ~1.6 ms and ~4.6 ms in both legs (periodic ~3 ms stalls visible in
the raw samples of every route), so the minima are the stable comparison.
A first version of the XLSX adapter counted the sheet's charts for a
default name on every add and measured 1.29× (N = 10) and 1.43× (N = 100);
it now counts only for an unnamed, untitled chart. Source-backed owners are
not measured: the PPTX unified add promotes a source owner to an eager
package first (one `to_bytes` and parse, as picture add does), which the
native API has no counterpart for.

## 6. Gaps and follow-ups

| Gap | Follow-up |
|---|---|
| DOCX charts (`add /body/p[N] --type chart`, 14 corpus lines): the leaf has the chart part graph but no inline-drawing authoring for `c:chart` graphic data | 0270a: `document::Edit` inline chart insertion next to `insert_picture_at`, reusing `chart::authoring::Spec` and `Workbook::from_spec` |
| Extended charts (`cx:chart`: histogram, box-whisker, funnel, treemap, pareto, sunburst) and waterfall | needs a chartex writer in `litchi-drawingml`; waterfall as OfficeCLI's stacked-bar simulation would read back as `stackedColumn`, so it is refused rather than mislabeled |
| Formatting keys not authored: fonts, gradients, transparency, fills, outlines, shadows, presets, trendlines, error bars, secondary axis, reference lines, data table, label rotation, tick marks, log scale, display units, bubble scale | extend `Appearance`; each is a write-only registry row today absent (`unsupported_property`) |
| Formatting keys are add-only and not read back by the node view | read them from `Summary` when a consumer needs them |
| `set` of chart type, data (`data=`), geometry (`x/y/width/height`) and frame name | type and whole-data replacement mean rebuilding the part; geometry set of graphic frames is a PPTX shape-write gap |
| `series{N}.name`/`.values` dotted keys, per-series `color` on `set`, axis `min`/`max`/`visible`/`format` on `set` | the source-preserving editor has `set_axis_range`; wire it to `axis` `min`/`max` |
| XLSX chart geometry is not read back (anchor not in the summary) | expose the anchor from the leaf inventory |
| Query walks list worksheet charts only when the walk is scoped to charts | merge the chart phase into the worksheet walk |
| A source-backed PPTX `remove` keeps the chart part and workbook (exact source replay policy) | an explicit dependency sweep on publication |

## 7. Not claimed

- Rendering fidelity beyond LibreOffice opening and drawing the charts with
  their titles; PowerPoint and Excel were not run.
- Any OfficeCLI behavior for keys outside §2 or §4.
