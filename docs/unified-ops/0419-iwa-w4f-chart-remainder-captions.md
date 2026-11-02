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
