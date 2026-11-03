# 0419 — W4-F: the remaining chart rows, captions and titles

- Status: in progress on branch `uo2/iwa-w4f` (cut from trunk
  `feat/unified-ops-api` at `02ad1c3981`).
- Date: 2026-11-02.
- Plan: [0363](0363-iwa-monolith-exit-plan.md) (§0 wave table, §3.2 API
  shape, §4 parity tiers, §5 performance); the previous chart wave
  [0414](0414-iwa-w4c-chart-style-families.md) (§6 remaining rows, §8 gate
  list, §11–§14 patterns); owner decision D1 and the
  [macOS queue](0363-mac-queue.md).
- Scope: the host-only `chart.style` rows (Keynote 8, Numbers 7, Pages 7),
  `chart.legend_axis` (Numbers 6, Pages 8), `chart.text` (Numbers 3, Pages
  3), and the Numbers drawable caption and title rows (shapes, images,
  movies). Chart add, duplicate and remove (`litchi_iwa_style::chart_lifecycle`),
  Pages chart geometry, Numbers shape image fill and Pages `body_table`
  belong to W4-E (0418); tables to W4-T.
- `performance_claim`: every migrated route measured is cheaper than its
  host route under callgrind (per-slice tables below); none misses the 5%
  bar.

## 1. Slice 1: the Radar series style

6 rows flip (set and read on each format).

| Crate | What |
|---|---|
| `litchi-iwa-style` | `chart_style::chart::{RadarSeriesStyle, Property::RadarSeriesStyle, PropertyKind::RadarSeriesStyle}` (the host's `ChartRadarSeriesStyle`: `Fill`, `Stroke`, `FillAndStroke`, the native default). New `series::style::radar`: the host's `charts::radar_series_style` rules. A value is read through each visible series' effective style and its parent chain: the area fill (field 165), the area stroke (172), the derived-fill switch (188, zero or one) and, when the switch is on, a finite alpha multiplier in 0–1 (189); the combination must be one of the three native modes, and every series must agree. A write reads every series color first (the effective stroke's color, else a solid fill's), then rewrites in series order each effective style whose value differs, exactly as the host patches it (fill 165 always, stroke 172 at 4 pt solid or removed, switch 188, multiplier 189 = 0.15 or removed), and verifies the uniform read. Payload layer `effective`/`rewrites` (the caller refuses a style another chart uses); bounded workspace layer behind `chart_style::{read_chart, write_chart}`, which now take the visible series count (the scope is the chart's series graph, as for the series-style family). The host checks no chart-style exclusivity for this property, and neither does the leaf |
| `litchi-numbers` | `sheet::chart::style` reads and writes the property on the payload layer and publishes the changed series styles in place; a shared series style is `Refusal::SharedStyle`. `sheet_chart_series::chart_series` exposes the visible series count to the style family |
| `litchi-keynote`, `litchi-pages` | `chart::style::RadarSeriesStyle`; the chart-style families pass the target's visible series count to the engine |

Parity: the options cases gain `Fill` and `Stroke` on a Radar chart, each
set and reset (to `FillAndStroke`) after a host set, and the reads of every
state. Numbers and Pages are byte-identical (Tier A), Keynote Tier B (root
previews). Regression
`numbers_sheet_chart_radar_series_style_shared_series_style_refuses`: a
second chart sharing the first chart's first series style makes the write
refuse in the host and the leaf, with the source unchanged. The existing
W4-C style rows of `0363-mac-queue.md` cover the routes.

| Route | Host Ir | Leaf Ir | Leaf/host |
|---|---:|---:|---:|
| keynote_radar_set | 853,472,234 | 78,570,948 | 0.0921 |
| pages_radar_set | 1,260,865,125 | 48,173,482 | 0.0382 |
| numbers_radar_set | 221,155,218 | 45,323,880 | 0.2049 |

Probe `crates/litchi-iwa/tests/parity/chart_property_cost.rs` (the 0414
method: release, CPU 3, ABBA, inclusive Ir of `w4c_host_route` and
`w4c_leaf_route`). Receipt `results/0419/callgrind-radar.json`; raw captures
in `.coord/iwa-w4f/cg1/`.

### 1.1 Slice 1 gates

Branch with trunk `56df8948b0` merged at `3c67e8b134`; logs
`.coord/iwa-w4f/gates1`. Every command of the 0414 §8 list exits 0: fmt;
clippy -D warnings on the five touched crates with `litchi-ops`, `litchi`
and `litchi-py` (all features), the leaf crates without features, and
`litchi-ops` with no features and with each of `numbers`, `pages`,
`keynote`; rustdoc -Dwarnings; the glossary check; the leaf, `litchi-iwa`
(parity included), `litchi`, `litchi-ops` (`--lib`, `--doc`, nine `--test`
groups) tests; pyo3; perf-baseline; the feature matrix; crate boundaries;
non-iWork gate; the tool unit tests; ledger, OfficeCLI delta and corpus
checks; the bench matrix (regenerated for 0419). No dispatch, capability or
corpus-visible behaviour changed, so no golden was re-blessed.

Trunk `ee40f1a375` (W4-E slice 1, W4-T 0412e) was then merged at
`befda82d89` (bench-matrix conflict resolved by regeneration). On that merge
fmt, clippy -D warnings on the five touched crates, the parity chart cases
(328 passed), the crate-boundary check, the glossary, the ledger and the
bench matrix exit 0 (logs `.coord/iwa-w4f/gates1b`). Target 13 GB after
pruning.

## 2. Slice 2: reference lines

6 rows flip (set and read on each format).

| Crate | What |
|---|---|
| `litchi-iwa-style` | `reference_line::{read_lines, write_lines}`: the W1-E graph engine (`lines`, `set_lines`, already host-differential since 0375a) under a residual `series::Limits` profile, returning a `Report` whose scope names the chart, every retained, removed and created line object, the chart and metadata members, the chart, line (5031) and package-metadata message types, and `REFERENCE_LINES_FIELD` (10005, the chart archive's `TSCH.ChartReferenceLinesArchive`) among the chart fields; created lines are planned new and removed ones planned removed. `series::MutationScope::allow_member` |
| `litchi-numbers` | `sheet::chart::reference_lines` (`Package::sheet_chart_reference_lines`, `edit_sheet_chart_reference_lines` → `Edit::set(Vec<Line>)`, `apply_sheet_chart_reference_lines`): the legend-style shape, published through the chart-series workspace path |
| `litchi-keynote` | `chart::reference_lines` on the chart-property core. The locality check admits the chart archive's field 10005 beside the generated extension when the write's chart fields name it, and line objects (5031) among the rewritable types |
| `litchi-pages` | `chart::reference_lines` on the chart-property core. The locality check admits field 10005 of the chart message when the plan allows it, and a planned line object rewritten whole |

The families share their format's chart-series `Error`, `LimitKind` and
`Refusal`. More than five lines, a repeated kind or a line another chart
lists refuse as the host does (`Error::InvalidSource`).

Parity: `{numbers_sheet,keynote_slide,pages_body}_chart_reference_lines_set_parity`
run five sequences per format through the facade with the inverse replayed:
allocate two lines into an empty axis, patch one in place (name and value
label), append one, remove one, remove all. Allocations are Tier B (the
host draws random item UUIDs; compared with the UUIDs zeroed, as in 0375a);
Numbers and Pages in-place and removal writes are byte-identical; Keynote
is Tier B throughout (root previews). `…_reference_lines_read_parity` reads
every state on each format (Tier A). Six macOS-queue rows added.

| Route | Host Ir | Leaf Ir | Leaf/host |
|---|---:|---:|---:|
| keynote_reference_lines_set | 879,854,138 | 139,136,007 | 0.1581 |
| pages_reference_lines_set | 1,276,021,892 | 88,109,646 | 0.0691 |
| numbers_reference_lines_set | 214,628,424 | 115,862,233 | 0.5398 |

Receipt `results/0419/callgrind-reference-lines.json`; raw captures in
`.coord/iwa-w4f/cg2/`.

### 2.1 Rows already served by a leaf route

Three host-only rows needed only parity and a ledger flip:

- Pages `body_chart_legend_visible` / `set_body_chart_legend_visible`:
  `chart::style::Property::LegendVisible` (0414 slice 1) on
  `pages_body_chart_options_{read,set}_parity` (Tier A); the existing W4-C
  Pages style queue rows cover them.
- Keynote `slide_chart_catalog`: the host route already reopens the package
  through `litchi_keynote::Package::slide_chart_catalog`; the new
  `keynote_slide_chart_catalog_read_parity` compares both on every slide of
  the native chart fixture and of `basic.key` with a host-added chart (Tier
  A, one queue row).

### 2.2 Slice 2 gates

Branch with trunk `4fc8e11883` merged at `c22d1206c0`; logs
`.coord/iwa-w4f/gates2`: every command of the 0414 §8 list exits 0 (the
same list as §1.1), plus the Python gates (boundaries, non-iWork, tool unit
tests, ledger, OfficeCLI delta and corpus, bench matrix).
