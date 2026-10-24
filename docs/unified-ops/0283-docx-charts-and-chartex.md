# 0283 — DOCX charts, extended chart types and chart `set` coverage

Date: 2026-10-14. Branch `uo2/charts-docx`, cut from trunk
`feat/unified-ops-api` at `80e24f8745`. Follows record 0270 (typed chart
CRUD; follow-ups 0270a DOCX charts, extended types, formatting keys, `set`
coverage) and record 0271 (dotted series keys).

## 1. What landed

### Leaf: `litchi-drawingml`

- `chart::authoring::Kind` gains the extended families `Waterfall`,
  `Funnel`, `Treemap`, `Sunburst`, `Histogram`, `Pareto` and `BoxWhisker`
  (`Kind::is_extended`). They take the standard flat variant only;
  waterfall, funnel, treemap, sunburst and Pareto take one series.
  `Appearance` gains `subtotals` (a waterfall's total points; default: the
  last), `increase_color`/`decrease_color`/`total_color`, `bins`
  (`Bins`: count or size, interval side, overflow, underflow),
  `quartile_method` and `parent_labels`.
- `chart::chartex` (new): `Spec::to_xml` writes an extended spec as a
  deterministic `cx:chartSpace` part — one `cx:data` per series with cached
  levels and, with a `Binding::Sheet`, `cx:f` references into the layout of
  `Spec::sheet_layout`; series with `layoutId` `waterfall`, `funnel`,
  `treemap`, `sunburst`, `boxWhisker`, `clusteredColumn` (histogram) or a
  `clusteredColumn` plus `paretoLine` pair (Pareto, with a percentage axis);
  layout properties (subtotals, binning or aggregation, quartiles, parent
  labels); data labels; per-point fills (waterfall increase/decrease/total
  colors, or a palette); axes; legend. A waterfall is a real `cx` waterfall
  with subtotal points, not a stacked-bar imitation. Titles are written as
  `cx:rich` (LibreOffice reads it; it ignores `cx:txData` titles).
  `chartex::read` reads a part back into the spec that writes it;
  `style_xml`/`colors_xml` write deterministic chart-style and color-style
  companions; `is_chartex` sniffs the root.
- `chart::authoring::read(xml) -> Parsed` (`Parsed { spec, external_data }`)
  reads a `c:` or `cx:` chart part back into its spec, with its formatting
  where the spec has a field (labels, gap, overlap, hole, slice angle,
  styles, smoothing, explosion, markers, line width, axis range and major
  unit, gridline visibility, legend overlay, auto-title, 3D view, series and
  point colors, combo families) and its binding inferred from the first
  series' references (`Spec::of(&Chart)` for a typed chart).
  `Parsed::is_authored(xml)` says whether rebuilding from the spec keeps
  everything the chart model reads (exact bytes for `cx:`; equal
  reader→writer normalisations for `c:`, so a part a source-preserving edit
  retouched still counts).
- `Spec::workbook_members` (moved out of PPTX) lists the members of the
  embedded chart workbook; PPTX and DOCX zip them.
- `chart::edit::Edit::set_series_color` sets or clears a series
  fill (the line color of line, scatter, stock and unfilled radar series),
  keeping the rest of its shape properties.

### Leaves: PPTX, XLSX, DOCX

- PPTX `opened::Edit::add_chart` writes an extended spec as a `chartEx`
  part related by `…/2014/relationships/chartEx`, with its style and color
  companions and the embedded workbook, and a `cx:chart` graphic frame;
  the scene classifies such frames as charts; `chart`/`set_chart` and the
  slide `chart_xml` readers accept them (`set_chart` keeps a part's family).
  `set_shape_properties` bounds now also move and resize chart graphic
  frames (`p:graphicFrame/p:xfrm`).
- XLSX `chart::edit` stages extended parts with their companions, counts
  `cx:chart` frames, keeps a part's family on `Edit::set`, and gains
  `Edit::set_anchor` (move/resize a two-cell anchor) with the reads
  `Package::chart_anchor` and `SourceWorksheet::chart_anchor`. The drawing
  inventory no longer copies every XML event (§4).
- DOCX: `Paragraph::charts() -> [InlineChart]` (relationship, name,
  description, extent, inline/anchored, extended);
  `Package::chart_xml` and `source_backed::Package::chart_xml`; on the
  package edit, `insert_chart_at(paragraph, at, &NewChart)` next to
  `insert_picture_at` — the chart part (or `chartEx` part with companions),
  a deterministic embedded workbook (`chart_workbook`, from
  `Spec::workbook_members`), the relationship and the inline drawing run in
  one graph change and one document operation — plus `chart_parts`,
  `set_chart` (same family only), `set_chart_extent_at` and
  `remove_chart_at` (the drawing, and its run when it holds nothing else;
  the part stays, unreferenced, as for pictures).

### Unified layer: `litchi-ops`

- **DOCX charts** are a real element: `add` under a paragraph
  (`/body/p[N]`, also in cells and content controls) with the shared chart
  vocabulary plus `name`, `alt`, `width`, `height` (lengths; default
  15 cm × 10 cm, OfficeCLI's), returning `/body/p[N]/chart[M]`. Charts
  resolve in their paragraph and document-wide in document order, as
  OfficeCLI returns them: `/chart[N]`, also `/body/chart[N]`. `get`/`query`
  read the frame and the chart part; `set` and `remove` as below. Series
  and axis child nodes are a DOCX gap (§5).
- **Extended types** on PPTX, XLSX and DOCX: `chartType=waterfall|wf,
  funnel, treemap, sunburst, histogram, pareto, boxWhisker`, with the keys
  `increaseColor`, `decreaseColor`, `totalColor`, `binCount`, `binSize`,
  `intervalClosed`, `overflowBin`, `underflowBin`, `quartileMethod`,
  `parentLabelLayout`. A one-series family given OfficeCLI's
  `data=Start:1000,Revenue:500,...` (one single-value series per item)
  takes the items as points and their names as categories.
- **`set` coverage** (all three hosts): `chartType` (same part family; an
  ordinary↔extended change is `unsupported_capability` with "remove and
  add"), `data` (series replaced, categories kept; XLSX writes replaced data
  inline), frame geometry (PPTX `x/y/width/height`, XLSX cell
  `x/y/width/height` through `set_anchor`, DOCX `width/height`),
  `axisTitle`, axis `min`/`max` on `axis[...]` (the editor's
  `set_axis_range`, the other bound kept) and chart-level `axisMin`/`axisMax`,
  per-series `color` on `series[K]` and `series{N}.color` on the chart, and
  every formatting key. Titles, legend, names, values, categories, colors
  and ranges stay source-preserving edits; a family change, new data, a
  formatting key and every edit of an extended chart **rebuild** the part
  from its spec, refused (`unsupported_capability`) unless
  `Parsed::is_authored` holds — a chart with content its spec cannot carry
  is never simplified. PPTX and DOCX regenerate the embedded workbook from
  the rebuilt spec when data or names change.
- **Read-back**: chart formatting keys are readable (`get: true`) when the
  chart sets them (`comboSplit` stays write-only); series nodes read
  `color`; axis nodes read `min`/`max`; XLSX charts read their anchor as
  `x/y/width/height` cells. The node views print them.
- Add admission's duplicate-key set grows to 256 entries (`Assigned`).
- `capabilities()`: DOCX `chart` is get/query/add/set/remove; registry and
  capability goldens re-blessed (§3).

## 2. Evidence

| Check | Result |
|---|---|
| `litchi-drawingml` unit tests | every extended family × binding round-trips through `chartex::read` and is authored; formatting (waterfall colors and subtotals, bins, quartiles, parent labels, axis range, legend) survives; invalid specs refused; foreign parts read but are not authored; companions parse as chart styles; every classic variant reads back as its authored spec, formatting reads back, a foreign `roundedCorners` is not authored; series colors insert, replace (ARGB), clear and keep opaque XML |
| PPTX unit test | extended chart: deterministic package, chartex + style + colors parts, a chart frame in the scene, part replaced in place, a classic part refused for it, removal sweeps the parts |
| XLSX unit test | classic and extended charts on one sheet, companions, `chart_anchor` on the eager and source owners, family kept on `set`, `set_anchor` no-op/move/refusal, removal sweeps companions |
| DOCX unit tests | classic and extended insert, staged reads, determinism, reopen, `Paragraph::charts`, eager and source `chart_xml`, `set_chart` (no-op, replace, family refusal), extent set/no-op/refusal, removal keeping the paragraph text, refused inputs stage nothing |
| `crates/litchi-ops/tests/chart_types.rs` (new) | DOCX round trip through `/body/p[2]/chart[1]`, `/chart[1]`, `/body/chart[1]` (source and eager), set of title + type + data + width + gapWidth + series color with workbook regeneration, extended rebuild, family refusal leaving bytes unchanged, exact no-op, remove, query; determinism and refusals; all seven extended types on PPTX, XLSX and DOCX (add, read, set); histogram/box/treemap keys read back and set; PPTX and XLSX `set` of chartType, data, x, axisMax, axis min, series color, axisTitle |
| `charts.rs` (0270/0271) | updated: `pieOfPie` stays unsupported, waterfall with two multi-point series is `invalid_value`, `name` not settable, `series1.color` settable, capability flags |
| LibreOffice 26.2.6.3 headless ([`results/0283/lo-check.json`](results/0283/lo-check.json), script [`lo_check.sh`](results/0283/lo_check.sh)) | DOCX: all 15 classic types render their titles. PPTX and XLSX: all 7 extended types render their titles and plotted values — LibreOffice draws them as ordinary charts (its ChartEx import has no waterfall/funnel/treemap/sunburst/histogram/Pareto/box shapes). DOCX extended charts: Writer does not import an inline ChartEx frame (no title rendered) |

## 3. OfficeCLI corpus

`LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test
officecli_corpus`, then `python3 tools/officecli_corpus.py` (and
`--check`). 6,904 rows, identical order; 130 rows changed, **no call that
passed regressed**.

| | Passed | Warned | Gap | Delta | Cascade | Bug |
|---|---:|---:|---:|---:|---:|---:|
| before (trunk `80e24f8745`) | 3,803 | 1,105 | 891 | 168 | 464 | 0 |
| this record | 3,847 | 1,172 | 779 | 169 | 464 | 0 |

| Chart `add` lines | Before | Now |
|---|---|---|
| DOCX (14) | 0 (`unsupported_element`) | 6 ok, 8 warned |
| PPTX extended (29) | `unsupported_capability` | 13 ok, 16 warned |
| XLSX extended (69) | `unsupported_capability` | 12 ok, 56 warned, 1 `invalid_argument` (D10) |

Row changes: `unsupported_capability`→warned 72, →ok 25, warned→ok 13,
`unsupported_element`→warned 8, →ok 6, `unsupported_capability`→
`invalid_argument` 1. Warned lines skip keys this record does not author
(fonts, fills, shadows, `preset`, …). New delta **D10**: a property given
twice in one chart `add` (a shared `DASH` option array sets
`dataLabels=false`, the line sets `dataLabels=true`) — OfficeCLI keeps the
last value, Litchi refuses with `invalid_argument` so no value is silently
dropped (`excel/charts/charts-histogram.sh:460`). The remaining 3 XLSX
`unsupported_capability` lines are the cell-range series of record 0271.

Goldens re-blessed with `--all-features` and with every `litchi-ops`
feature except `encryption`: `capabilities/{docx,docx-encryption,pptx,
xlsx,xlsx-encryption}.json` and `registry/{docx,docx-encryption,pptx,xlsx,
xlsx-encryption}.json` (chart, series and axis entries only, and DOCX
paragraph children gain `chart`); `officecli-capability-delta.md`
regenerated.

## 4. Measurements

Harness `tools/perf-baseline/unified-chart-add` (README there), extended
with `--formats docx` (route `docx_chart_add_v1`) and `--compare set`
(routes `{pptx,xlsx,docx}_chart_set_v1`: the native chart retitle — the
source-preserving editor's `set_title` plus the host's part replacement and
one commit — versus `File::set(<chart[1]>, {title})`). Release build of
commit `10de671316`
(`CARGO_NET_OFFLINE=true cargo build --release --manifest-path
tools/perf-baseline/unified-chart-add/Cargo.toml`, binary SHA-256
`c08e81b35406bfdbfec7dbd2edab64201aa387bf73cbe55af1b0b2eb272a6b05`, not
committed), `taskset -c 13`, 15 ABBA rounds (30 samples per leg), two
trials, load average 8.7–9.1 on 32 cores. Both legs publish exact bytes
(PPTX add: equal decoded members, as in record 0270). Raw:
[`results/0283/add-{1,2}.json`](results/0283/),
[`results/0283/set-{1,2}.json`](results/0283/).

| Route | N | Native median | Unified median | Unified / native (t1 / t2) |
|---|---:|---:|---:|---|
| `pptx_chart_add_v1` | 0 / 10 / 100 | 0.89 / 1.54 / 6.40 ms | 0.90 / 1.55 / 6.40 ms | 1.016 / 1.014; 1.010 / 1.015; 1.001 / 1.002 |
| `xlsx_chart_add_v1` | 0 / 10 / 100 | 46.4 / 108.5 / 663.7 µs | 49.3 / 111.6 / 666.5 µs | 1.062 / 1.059; 1.029 / 1.027; 1.004 / 1.008 |
| `docx_chart_add_v1` | 0 / 10 / 100 | 99.5 / 185.2 / 940.1 µs | 105.6 / 191.1 / 967.5 µs | 1.061 / 1.068; 1.032 / 1.058; 1.029 / 1.050 |
| `pptx_chart_set_v1` | 1 / 10 / 100 | 0.64 / 0.79 / 2.18 ms | 0.65 / 0.80 / 2.18 ms | 1.014 / 1.007; 1.015 / 1.015; 0.999 / 1.004 |
| `xlsx_chart_set_v1` | 1 / 10 / 100 | 178.3 / 229.5 / 729.6 µs | 184.7 / 259.1 / 985.8 µs | 1.036 / 1.035; **1.129 / 1.121; 1.351 / 1.348** |
| `docx_chart_set_v1` | 1 / 10 / 100 | 216.3 / 252.8 / 571.9 µs | 220.4 / 259.5 / 590.1 µs | 1.019 / 1.020; 1.027 / 1.026; 1.032 / 1.034 |

Every add and every PPTX and DOCX set is within the 10% single-op budget.
**XLSX chart `set` misses it at N = 10 (1.12–1.13) and N = 100 (1.35)**:
the unified call resolves `/Sheet1/chart[1]`, which reads the worksheet
drawing's chart inventory once more than the native edit (which reads it
twice, in `Edit::chart` and `Edit::set`); the inventory of a 100-anchor
drawing costs about 240 µs (measured through `Package::chart_count`). Two
earlier versions measured worse and were fixed here: an eager chart-part
projection and authoring-spec build on every resolution (DOCX 1.18–1.22,
XLSX 1.28–1.55), now read only when a key needs them, and a chart count
plus an anchor read on every positional step, now skipped; the inventory
itself no longer copies each XML event (−20%). Follow-up below.

## 5. Gaps and follow-ups

| Gap | Follow-up |
|---|---|
| XLSX chart `set` 1.12–1.35× at 10–100 charts (one extra drawing inventory in path resolution) | memoize the drawing's chart inventory per worksheet snapshot in `litchi-xlsx` (shared by resolution, `Edit::chart` and `Edit::set`), or a faster inventory scan |
| DOCX chart series and axis child nodes (`/chart[N]/series[K]`, `axis[...]`); DOCX series color and axis range go through `series{N}.color` and `axisMin`/`axisMax` on the chart | add `Series`/`Axis` kinds to the DOCX vocabulary (as PPTX and XLSX have them) |
| DOCX extended charts are not drawn by LibreOffice Writer; Word and Excel were not run | an `mc:AlternateContent` fallback picture would let other readers show something; verify in Word |
| LibreOffice draws PPTX/XLSX extended charts as ordinary charts | upstream LibreOffice ChartEx import |
| XLSX extended charts with inline `data` carry cached levels without `cx:f` (no host cells written); Excel's acceptance not verified | write the data to host cells as OfficeCLI does, or bind with `dataRange` |
| Rebuild-based `set` refuses a chart with content its spec cannot carry (fonts, fills, trendlines, …) | extend `Appearance` with those keys (record 0270 gap 3) |
| `chartType` between an ordinary and an extended family | swap the part (content type, relationship, frame URI) in each host |
| Formatting keys still unauthored: fonts, gradients, transparency, chart/plot fills, outlines, shadows, `preset`, trendlines, error bars, secondary axis, reference lines, data table, label rotation, display units | record 0270 gap 3 |
| A removed DOCX chart keeps its part and relationship (unreferenced) | a dependency sweep on publication, as for pictures |
| `name`/`alt` of a DOCX chart frame are add-only | a `docPr` attribute edit in the DOCX transaction |

## 6. Gates

Run with `CARGO_TARGET_DIR=<worktree>/target CARGO_INCREMENTAL=0
CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0`.

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-drawingml -p litchi-pptx -p litchi-xlsx -p litchi-docx -p litchi-spreadsheet-drawing -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features docx\|pptx\|xlsx] --all-targets -- -D warnings` (none, docx, pptx, xlsx) | 0 each |
| `cargo test -p litchi-drawingml --all-features` (226 lib + doc tests) | 0 |
| `cargo test -p litchi-pptx --all-features`, `-p litchi-xlsx`, `-p litchi-docx`, `-p litchi-spreadsheet-drawing` | 0 each |
| `cargo test -p litchi-ops --all-features` (2,108 passed) | 0 |
| `cargo test -p litchi --all-features` (260 passed) | 0 |
| `cargo test -p litchi-cli --all-features` (corpus included) | 0 |
| capability/registry goldens with `--all-features` and with every feature but `encryption` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-drawingml -p litchi-pptx -p litchi-xlsx -p litchi-docx -p litchi-spreadsheet-drawing -p litchi-ops` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (26 profiles, 1,201 bans, 39 families) | 0 |
| `python3 tools/officecli_corpus.py --check` | 0 |
| `python3 tools/officecli_capability_delta.py --officecli <OfficeCLI> --check` | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_no_3rdparty_references tools.test_officecli_capability_delta` | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |

Python bindings are unchanged (they compile under the clippy gate); the
`pyo3-litchi` suite was not run.

## 7. Not claimed

- Rendering of extended charts, chart colors or rebuilt charts in Word,
  PowerPoint or Excel; LibreOffice behaviour is as in §2.
- That Excel opens XLSX extended charts authored from inline data (§5).
- Any OfficeCLI behaviour for keys outside §1.
