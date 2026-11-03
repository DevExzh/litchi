# 0425 — W4-L: series and symbol image fills, the shared caption lift, Numbers and chart captions

- Status: slice 1 merged (`bab7c866b4`); slices 2 (caption lift) and 3
  (Numbers drawable captions) merge-ready; stopped at a clean commit with a
  successor brief (§4) for the Pages chart captions and the Numbers
  componentized chart caption. Branch `uo2/iwa-w4l`, cut from trunk
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

### 1.1 Slice 1 gates

Branch with trunk `a077aeeb43` (W4-J 0423) merged at `0049a606e5`
(mac-queue conflict resolved by keeping both sides, bench matrix
regenerated); logs `.coord/iwa-w4l/gates1`. Every command of the 0414 §8
list exits 0: fmt; the seven Python gates (boundaries, non-iWork, ledger,
OfficeCLI delta and corpus, bench matrix, tool unit tests); clippy
-D warnings with all features, the leaf crates without features, and
`litchi-ops` with no features and each of `numbers`, `pages`, `keynote`;
the leaf, `litchi-iwa` (parity included), `litchi` and `litchi-ops`
(`--lib`, `--doc`, nine `--test` groups) tests; pyo3; perf-baseline; the
feature matrix. Rustdoc -Dwarnings and the glossary first failed on two
ambiguous `write` links in `series::style` (the function and the `write!`
macro); after the fix both exit 0 (`rustdoc2.log`, `glossary2.log`). No
dispatch, capability or corpus-visible behaviour changed, so no golden was
re-blessed. Target 21 GB after the gate run (test executables and examples included), 12 GB after pruning.

## 2. Slice 2: the shared caption lift (`f5392887fd`, own commit)

No ledger row; parity unchanged.

| Crate | What |
|---|---|
| `litchi-iwa-style` | New `caption`: the inline title/caption graph of a drawable. Payload- and object-level rules: `Label` (fields 10/11, native kind), `Nesting` (drawable at field 1, or 1.1 for shapes), `GraphIds`, `Theme`, `graph_objects` (the four objects over `keynote_chart_caption_graph_codec`, host versions and references), `standin`, `info_graph` (caption-info validation: kind, parent, owned and deprecated storage, text box, distinct objects), `reference_payload`, `repoint`, `replace_reference` (the host's `replace_object_reference`), `label_reference`, `drawable_width`/`archive_width`, `theme_paragraph_style` (`TSS.ThemeArchive` presets 210.1). Workspace layer for owners that publish workspace changes: `resolve`, `read`, `write` (text replacement through the TSWP storage rewrite; a new inline graph with identifiers from the watermark, UUIDs in the drawable's component and the watermark raised; removal to a fresh stand-in, the old graph kept) and `component_holds` |
| `litchi-pages` | `body_drawable_caption` builds labels on it; its `Kind` field/graph mapping, caption-info validation, width reader, theme preset parse, `reference_payload`, `repoint`, `replace_reference`, `versioned_object` and graph builder are deleted (303 → net −190 lines) |
| `litchi-keynote` | The caption owner (`slide_chart_caption`, serving chart, movie, image and shape captions) takes its graph objects, stand-in, identifier allocation and drawable width from it; its copies are deleted. Keynote's strict theme-reference parse and its bounded publication and verification seams are unchanged |

Checks: the 30 Pages and Keynote caption/title parity cases and the leaf
caption tests pass unchanged; three unit tests cover the shared reference,
identifier and edge rules.

## 3. Slice 3: Numbers drawable titles and captions (`c9c961dd65`)

17 rows flip: `sheet_{shape,image,movie}_title_caption`,
`{set,remove}_sheet_{shape,image,movie}_{title,caption}`, and the chart
caption read and removal (`sheet_chart_caption`,
`remove_sheet_chart_caption`).

| Crate | What |
|---|---|
| `litchi-numbers` | `sheet::drawable::caption::{Edit, Commit, Patch, Diagnostics, Error, LimitKind, Refusal, Label}`; `Package::{sheet_drawable_caption, edit_sheet_drawable_caption, apply_sheet_drawable_caption}` (`Edit::{set, clear, before}`) for shapes, images, movies and charts (caption only) on `caption::write` over the style workspace, published by `style_graph::publish`; the caption theme comes from the document (stylesheet 4, theme 6, language 8.3) and `TN.ThemeArchive.super`. Text boxes, audio, tables and chart titles are `Refusal::NoLabel`; a new caption of a chart stored beside the calculation engine is `Refusal::ComponentizedLayout` |

Parity (`numbers_caption_cases.rs`): set (new graph, then text
replacement after host labels) and remove for each kind and label, Tier B
(new object UUIDs); `numbers_sheet_drawable_caption_read_parity` (Tier A)
on unlabelled and host-labelled drawables; chart caption replacement and
removal; `numbers_sheet_chart_caption_componentized_refuses` (stricter
admission, macOS-queued). Two macOS-queue rows.

| Route | Host Ir | Leaf Ir | Leaf/host |
|---|---:|---:|---:|
| shape_title_create | 345,663,742 | 88,374,469 | 0.2557 |
| image_caption_create | 276,665,634 | 94,582,727 | 0.3419 |
| movie_caption_replace | 237,703,642 | 57,946,532 | 0.2438 |
| image_title_remove | 270,770,781 | 94,340,639 | 0.3484 |
| chart_caption_remove | 283,358,265 | 84,107,402 | 0.2968 |

Probe `numbers_caption_route_cost_abba` (`W4L_ROUTE`, release, CPU 3,
ABBA, inclusive Ir of `w4l_host_route`/`w4l_leaf_route`). Receipt
`results/0425/callgrind-numbers-captions.json`; raw captures in
`.coord/iwa-w4l/cg3/`.

## 4. Successor brief

Still host-only in this wave's scope (4 rows):

1. **Numbers `set_sheet_chart_caption`.** Numbers charts live in the
   document component beside the calculation engine, where the host
   authors the *componentized* caption graph
   (`image_caption.rs`, `CaptionGraphProfile::Componentized`;
   `numbers/editor/sheet_charts/caption.rs`): the label's shape style goes
   into `Index/DocumentStylesheet.iwa` and is registered in the stylesheet
   (401) object's references, the placement payload differs, only
   `component_objects()` get UUIDs, and the chart message's reference list
   keeps the old stand-in and appends the info. The failing diff is
   reproducible by deleting the `ComponentizedLayout` refusal in
   `sheet_drawable_caption.rs` and running `numbers_sheet_chart_caption_*`.
   Add a `caption::Layout::Componentized` to the shared `graph_objects`/
   `write` (the codec may need a placement variant) and drop the refusal.
2. **Pages chart captions** (`body_chart_caption`, `set_body_chart_caption`,
   `remove_body_chart_caption`, `pages/editor/charts/caption.rs`). Either
   add a chart path to `body_drawable_caption` (a `BodyChartSelector`
   entry resolving the chart drawable, `Nesting::Drawable`, caption only)
   using the shared payload rules, or a workspace route on the chart core.
   Check first which layout the host uses for Pages charts.

Patterns: Numbers families on `style_graph::publish` take the
`sheet_drawable_style.rs` scaffolding (`structural_family!`, `workspace`,
`Commit::new`); the ledger tool reads parity test names literally, so do
not generate tests with a macro.
