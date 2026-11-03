# 0425 — W4-L: series and symbol image fills, the shared caption lift, Numbers and chart captions

- Status: slice 1 merge-ready. Branch `uo2/iwa-w4l`, cut from trunk
  `feat/unified-ops-api` at `0ea106b34c`.
- Date: 2026-11-03.
- Plan: [0363](0363-iwa-monolith-exit-plan.md) (§0, §3.2, §4, §8); the
  predecessor's successor brief [0422](0422-iwa-w4i-value-axis-image-fills-captions.md)
  §5; the chart patterns and gate list of
  [0414](0414-iwa-w4c-chart-style-families.md) (§8, §11–§14); owner
  decision D1 and the [macOS queue](0363-mac-queue.md).
- Scope, in order: series and symbol image fills (6 rows); the shared
  caption lift (own commit, parity unchanged); Numbers shape, image and
  movie captions and titles (15 rows); Pages and Numbers chart captions
  (`chart.text`, 6 rows).
- `performance_claim`: every migrated route measured is cheaper than its
  host route under callgrind (per-slice tables below); none misses the 5%
  bar.

## 1. Slice 1: series and data-symbol image fills

6 rows flip (`set_{body,sheet,slide}_chart_series_image_fill` and
`set_*_chart_series_symbol_image_fill`).

| Crate | What |
|---|---|
| `litchi-iwa-style` | `series::style::write_data` (+ `SeriesImage`): the series-style write with the host's data-owner transitions for fills and data-symbol fills (`set_chart_series_{,symbol_}fills`, their resets, `set_chart_series_{,symbol_}image_fill_data`). A new image is registered first (`data_owner::own_new`, the series' effective style as owner, after its exclusivity check: the host's `insert_unreferenced` then `add_component_data_reference`) and becomes that series' value; an image value naming stored data adds the style as one more owner (new `data_owner::own`, additive; the rest of `data_owner` is unchanged); a changed series whose own fill (for symbols: the fill message whatever its switches) held an image releases it (`data_owner::release`), slot by slot as the host's `adjust_data_reference` + `remove_orphaned_image_asset`. `series::style::write` now refuses a write that would change data owners (`Refusal::ImageFill`) instead of leaving a stale owner. `chart_style::DataChanges::deleted` is a list (a set-all may release several images); the scope planning of `write_chart_background` is the shared `chart_style::plan_data_changes` |
| `litchi-pages`, `litchi-keynote` | `chart::series::Edit::{set_image_fill, set_symbol_image_fill}(series, bytes, file_name, ImageTechnique, tint)` (frame size: Pages `body_chart_geometry`, Keynote `chart_style::read_frame_size`); fill and symbol-fill `set`/`reset` go through `Ledger::write_data`, so a replaced image is released and its member deleted. `series::{ImageTechnique, Rgba}` re-exported |
| `litchi-numbers` | The same verbs; fill and symbol-fill writes move from the payload-level in-place path to the chart-series workspace path (`write_data` + `publish_changes_with_data`, as the background image) |

Parity (`chart_series_image_fill_cases.rs`):
`{pages_body,numbers_sheet,keynote_slide}_chart_series_image_fill_parity`
(column chart) and `…_series_symbol_image_fill_parity` (line chart) run,
on series 1, a new image (scaled; tiled and tinted), a new image over a
host image (the old one released and its `Data/` member deleted), and a
solid fill and a reset over a host image; the symbol cases add
follow-the-series-fill over a host image. Each replays the inverse and
compares the host's fill reads and the leaf's fill, symbol-fill and media
reads. Tier B (ZIP layout of the new member; Keynote also root previews),
macOS-queued (six rows). The existing series-style, background-image and
fill parity cases (210) still pass.

| Route | Host Ir | Leaf Ir | Leaf/host |
|---|---:|---:|---:|
| pages_series_image_set | 1,384,519,495 | 64,002,177 | 0.0462 |
| numbers_series_image_set | 359,702,959 | 89,000,789 | 0.2474 |
| keynote_series_image_set | 1,062,506,049 | 99,823,794 | 0.0940 |
| pages_symbol_image_set | 1,384,539,665 | 63,897,464 | 0.0462 |
| numbers_symbol_image_set | 359,798,713 | 88,880,670 | 0.2470 |
| keynote_symbol_image_set | 1,062,509,608 | 99,787,842 | 0.0939 |

Probe `chart_property_cost.rs` (`W4C_ROUTE`, release, CPU 3, ABBA,
inclusive Ir of `w4c_host_route`/`w4c_leaf_route`; series 1 gets a tiled
image). Receipt `results/0425/callgrind-series-image-fills.json`; raw
captures in `.coord/iwa-w4l/cg1/` (outside the repository). The bench
matrix lists 0425 under `evidence_without_rows` (leaf-versus-host
evidence).
