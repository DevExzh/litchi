# 0422 — W4-I: value-axis settings, chart image fills, captions and titles

- Status: slices 1 and 2 merged (`41ee75653b`); slice 3 merge-ready; stopped at a clean commit (context) with a successor brief (§5). Branch `uo2/iwa-w4i`
  (cut from trunk `feat/unified-ops-api` at `c599ef9cb8`).
- Date: 2026-11-02.
- Plan: [0363](0363-iwa-monolith-exit-plan.md) (§0 wave table, §3.2 API
  shape, §4 parity tiers, §5 performance); the predecessor's successor brief
  [0419](0419-iwa-w4f-chart-remainder-captions.md) §3; the chart patterns
  and gate list of [0414](0414-iwa-w4c-chart-style-families.md) (§8,
  §11–§14); owner decision D1 and the [macOS queue](0363-mac-queue.md).
- Scope, in order: value-axis bounds, steps and scale (`chart.legend_axis`,
  Pages 6, Numbers 6); chart image fills (background, series, symbol; 3 per
  format); the shared caption lift, then the Numbers shape, image and movie
  captions and titles (15) and the chart captions (Pages 3, Numbers 3).
  Chart add/duplicate/remove (`chart_lifecycle`) and `table_create` belong to
  W4-G; table sort and pivot to W4-H.
- `performance_claim`: every migrated route measured is cheaper than its
  host route under callgrind (per-slice tables below); none misses the 5%
  bar.

## 1. Slice 1: value-axis bounds, steps and scale

12 rows flip (Pages and Numbers: the three host setters and the three host
reads each).

| Crate | What |
|---|---|
| `litchi-iwa-style` | New `value_axis`: Keynote's rules, lifted. `validate` (finite endpoints, minimum ≤ maximum, no non-positive endpoint on a logarithmic axis; a known scale given by its native value is canonicalised), `equal_exact`, `from_snapshot` (a stored aggregate the native format could not have written is invalid), `codec_write` (presence-preserving: a component equal to the stored one keeps its bytes, an absent scale field stays absent while linear) over `litchi_iwa_protos::keynote_chart_axis_value_settings_codec`. Payload layer `settings`/`rewrite` on a `TSCH.ChartAxisNonStyleArchive` payload: the rewrite applies the components in the host's order (bounds, steps, scale; minimum before maximum, major before minor), so fields absent from the source are appended exactly where the host's `set_chart_value_axis_{bounds,steps,scale}` append them. Workspace layer `read`/`write`: the chart's primary value-axis non-style object, refused when another chart axis references it (the host's `AxisNonStyleSlot::ensure_exclusive`), rewritten in place under a residual `series::Limits` profile with the in-place `MutationScope` (extension field 10,000 of the axis object). New engine `Refusal::InvalidSettings`; `chart_style::{owners, style_object, in_place_scope, rewrite_in_place, read_payload}` are crate-visible for it |
| `litchi-keynote` | `slide_chart_value_axis` keeps its bounded transaction (budget, selection, locality) and now calls the shared rules; its private `common_settings`, `validate_settings`, `settings_equal_exact`, `codec_settings` and `codec_write` (124 lines) are deleted. Behaviour unchanged (`tests/slide_chart_value_axis.rs`, `tests/chart_axis_api.rs`) |
| `litchi-pages` | `chart::value_axis` (`Package::body_chart_value_axis_settings`, `edit_body_chart_value_axis_settings` → `Edit::{set, set_bounds, set_steps, set_scale}`, `apply_body_chart_value_axis_settings`) on the chart-property core, Keynote's aggregate shape. Chart-series `Refusal::InvalidSettings` (`RefusalKind::InvalidValue`) |
| `litchi-numbers` | `sheet::chart::value_axis`, the same family on the chart-series workspace path (`&mut` staging, as the other Numbers chart families). Chart-series `Refusal::InvalidSettings` |

Parity (`value_axis_cases.rs`): `{pages_body,numbers_sheet}_chart_value_axis_set_parity`
run 13 sequences per format through the facade with the inverse replayed:
bounds, steps and scale into an axis that does not store them, a single
maximum, in-place bounds and steps, a scale over stored bounds, back to
automatic bounds and steps, logarithmic back to linear (the stored scale
stays explicit), both step fields removed and appended again, and two
three-component edits (the host writes bounds, steps, scale in order). All
Tier A (byte-identical, whole ZIP). `…_read_parity` compares the host's three
reads with the aggregate read on every state (Tier A). Regressions:

- `numbers_sheet_chart_value_axis_shared_axis_refuses`: a second chart whose
  value axis references the first chart's non-style object; the host and the
  leaf refuse, the source is unchanged.
- `chart_value_axis_logarithmic_non_positive_bound_refuses` (stricter
  admission, §4.2 of 0363): the host stores a logarithmic scale over a
  negative minimum; the leaf refuses the staging (`Refusal::InvalidSettings`)
  as Keynote always has, and reads such a source as invalid. Listed in the
  macOS queue.

Eight macOS-queue rows added (the six setters and two reads).

| Route | Host Ir | Leaf Ir | Leaf/host |
|---|---:|---:|---:|
| pages_value_axis_set | 1,193,208,991 | 48,131,065 | 0.0403 |
| numbers_value_axis_set | 151,334,720 | 77,925,554 | 0.5149 |

Probe `chart_property_cost.rs` (`W4C_ROUTE`, release, CPU 3, ABBA, inclusive
Ir of `w4c_host_route`/`w4c_leaf_route`; the route sets fixed bounds 0–50).
Receipt `results/0422/callgrind-value-axis.json`; raw captures in
`.coord/iwa-w4i/cg1/`. Keynote has no host value-axis route to compare; its
transaction path is unchanged apart from the shared rule calls.

## 2. Slice 2: shared data-owner transitions; Pages chart background image

1 row flips (Pages `set_body_chart_background_image_fill`); the existing
Pages background-fill set also gains the host's release of a replaced image.

| Crate | What |
|---|---|
| `litchi-iwa-style` | New `data_owner`: the `TSP.PackageMetadata` payload helpers lifted from Pages (`adjust_data_reference`, `referenced_data`, `data_member`, `remove_data_info`, `data_metadata_entries`, `data_owners`, `max_data_identifier`, `append_data_info`), `materialized_name`, and two workspace transitions on the registry: `own_new` (next data identifier, `DataInfo`, the object as its component data-reference owner: the host's `insert_unreferenced` + `add_component_data_reference`) and `release` (drop the owner; data no component, data metadata map entry or message-info reference keeps loses its `DataInfo` and member: the host's `remove_component_data_reference` + `remove_orphaned_image_asset`). `chart_style::write_chart_background(…, Background::{Fill, Image(NewImage)})` stores a background with these transitions (exclusive chart style; new image owned by the chart style; a replaced image released) and reports `DataChanges` (inserted and deleted `Data/` members), with the metadata object, fields 3/4/7/11 and the planned members in its mutation scope. `chart::rewrite_background` is the image-admitting payload rule; `chart::rewrite` still refuses an image (`Refusal::ImageFill`) |
| `litchi-pages` | `body_storage_objects` delegates its metadata helpers to `data_owner` (the copies are deleted). The chart-property core gains `Ledger::write_data` (publishes the new `Data/` member and deletes the orphaned one); its locality check admits the metadata `DataInfo` list (4) and component data references (7) when the plan allows them. `chart::style::Edit::set_background_image_fill(bytes, file_name, ImageTechnique, tint)` (frame size from `body_chart_geometry`), and `Edit::set(BackgroundFill)` now goes through the releasing write. Chart-series `Refusal::InvalidImage` |

Parity `pages_body_chart_background_image_fill_parity` (Tier B, the ZIP
layout of the new member; macOS-queued): a new image (scaled, tiled and
tinted), a new image over a host image (the old one is released and its
`Data/` member deleted), and a solid fill and no fill over a host image;
each with the inverse replayed and the fill and media reads compared. The
host keeps an unreferenced asset only when the data metadata map lists it;
the leaf's `release` follows that rule (the first draft kept every asset of
a package with a map, which the parity case caught).

| Route | Host Ir | Leaf Ir | Leaf/host |
|---|---:|---:|---:|
| pages_background_image_set | 1,310,073,205 | 63,127,004 | 0.0482 |

Receipt `results/0422/callgrind-background-image.json`; raw captures in
`.coord/iwa-w4i/cg2/`.

### 2.1 Gates (slices 1 and 2)

Branch with trunk `44aa1ee06d` (W4-G 0420) merged at `4ea803b0dc`
(mac-queue conflict resolved by keeping both sides); logs
`.coord/iwa-w4i/gates2`. Every command of the 0414 §8 list exits 0: fmt;
clippy -D warnings on the five touched crates with `litchi-ops`, `litchi`
and `litchi-py` (all features), the leaf crates without features, and
`litchi-ops` with no features and with each of `numbers`, `pages`,
`keynote`; the leaf, `litchi-iwa` (parity included), `litchi` and
`litchi-ops` (`--lib`, `--doc`, nine `--test` groups) tests; pyo3;
perf-baseline; the feature matrix; crate boundaries; non-iWork gate; the
tool unit tests; ledger, OfficeCLI delta and corpus checks. Rustdoc
-Dwarnings and the glossary first failed on one broken module doc link in
`data_owner` (fixed in `5302ab5ffc`) and then exit 0. The bench matrix
needed a 0422 `evidence_without_rows` entry (leaf-versus-host evidence, as
for 0419) and a regeneration (`c55a619959`); then exit 0. No dispatch,
capability or corpus-visible behaviour changed, so no golden was
re-blessed. Target 13 GB after pruning.

## 3. Slice 3: chart background images on Numbers and Keynote

2 rows flip (`set_sheet_chart_background_image_fill`,
`set_slide_chart_background_image_fill`); on both formats a background set
over a stored image now releases it, as on Pages.

| Crate | What |
|---|---|
| `litchi-iwa-style` | `chart_style::read_frame_size`: the chart drawable's frame size (the image fill's layout frame, the host's `graph.info.geometry.size`) |
| `litchi-numbers` | `sheet::chart::style::Edit::set_background_image_fill` and the image-releasing background set run on the chart-series workspace path (`write_chart_background`); `publish_changes_with_data` inserts the `Data/` member and the member-locality check counts it; frame size from `sheet_drawable_properties`. Chart-style and chart-series `Refusal::InvalidImage`; the chart-style family now maps the series family's `MissingStyle` refusal |
| `litchi-keynote` | `chart::style::Edit::set_background_image_fill` on the chart-property core: `Ledger::write_data` publishes the data member; the series-scope mutation plan lists the engine's planned new members; the metadata locality admits the `DataInfo` list (4) when the plan allows it. Chart-series `Refusal::InvalidImage` |

Parity `{numbers_sheet,keynote_slide}_chart_background_image_fill_parity`:
the Pages sequences (§2), Tier B (Keynote: root previews), macOS-queued.

| Route | Host Ir | Leaf Ir | Leaf/host |
|---|---:|---:|---:|
| numbers_background_image_set | 279,063,305 | 86,463,437 | 0.3098 |
| keynote_background_image_set | 962,539,495 | 99,001,064 | 0.1029 |

Receipt `results/0422/callgrind-background-image-numbers-keynote.json`; raw
captures in `.coord/iwa-w4i/cg3/`.

### 3.1 Slice 3 gates

Branch with trunk `41ee75653b` merged at `f1425d9c6f` (mac-queue conflict
resolved by keeping both sides); logs `.coord/iwa-w4i/gates3`: all 33 rows
exit 0 (fmt; the seven Python gates including the ledger and the bench
matrix; clippy -D warnings with all features, without features and per
`litchi-ops` format feature; rustdoc -Dwarnings; the glossary; the leaf,
`litchi-iwa` (parity included), `litchi` and `litchi-ops` tests in nine
groups; pyo3; perf-baseline; the feature matrix). No golden re-blessed.
Target 13 GB after pruning.

## 4. Remaining

- Series and symbol image fills (6 rows). Every core now publishes data
  members (`write_data`, `publish_changes_with_data`) and admits the
  metadata `DataInfo`/data-reference fields; what is missing is the engine:
  `series::style` refuses `Fill::Image` (`Refusal::ImageFill`). The host
  (`charts::series_fill::set_chart_series_image_fill_data`,
  `series_symbol_fill`) inserts the media, sets the whole fill vector and
  moves the data reference per series style slot
  (`adjust_data_reference`, then `remove_orphaned_image_asset`). The engine
  needs the same per-series owner transition on the (possibly newly
  allocated) series style variation, with `data_owner::{own_new, release}`
  and `DataChanges` as in `write_chart_background`.
- The shared caption lift, then Numbers shape, image and movie captions
  and titles (15 rows), then chart captions (Pages 3, Numbers 3).

## 5. Successor brief

Ledger after slice 3: W4-I has flipped 15 rows (value axis 12, background
images 3). Still host-only in this wave's scope: series and symbol image
fills (Keynote, Pages, Numbers: 6), Numbers shape, image and movie captions
and titles (15), Pages and Numbers chart captions (6).

1. **Series and symbol image fills** (`{body,sheet,slide}/…/series_fill.rs`
   and `series_symbol_fill.rs`, `set_*_chart_series_image_fill`,
   `set_*_chart_series_symbol_image_fill`). Publication is done on all three
   cores (`Ledger::write_data` on Pages and Keynote,
   `sheet_chart_series::publish_changes_with_data` on Numbers; the metadata
   `DataInfo` (4) and component data-reference (7) fields are admitted when
   the scope allows them). The work left is in `litchi_iwa_style::series`:
   `series::style` refuses `Fill::Image`. Follow the host's
   `charts::series_fill::set_chart_series_image_fill_data` (insert media,
   set the whole fill vector, `adjust_data_reference` per series style slot,
   then `remove_orphaned_image_asset`) with `data_owner::{own_new, release}`
   on the series style variation that ends up storing the fill (it may be
   newly allocated, so own the data after the allocation), return a
   `chart_style::DataChanges`, and add `metadata` + fields 3/4/7/11 and the
   planned `Data/` members to the scope exactly as `write_chart_background`
   does. Parity: copy `chart_image_fill_cases.rs` (the five sequences per
   format). The coordinator asked that the `data_owner` API stay stable
   (W4-J uses it for Numbers shape image fills).
2. **Shared caption lift** (its own commit): Keynote
   `slide_chart_caption.rs` (4,626 lines, record 0372) and
   `slide_movie_caption.rs` (2,295), Pages `body_drawable_caption.rs`
   (1,034; images, movies, shapes) into `litchi-iwa-style`, general for
   shapes, images, movies and chart drawables; the codecs
   (`keynote_chart_caption_graph_codec`, `pages_movie_caption_codec`,
   `keynote_movie_caption_codec`) already live in `litchi-iwa-protos`.
   Pages' module doc states the graph: title and caption are drawable
   fields 10 and 11; a present label is a `TSA.CaptionInfoArchive` with
   a private storage, shape style and placement in the drawable's
   component; clearing points the drawable at a fresh
   `TSD.StandinCaptionArchive`. Move Keynote and Pages onto it, delete
   their copies.
3. **Numbers drawable captions and titles** (15 rows:
   `numbers/editor/{sheet_shapes/caption.rs, sheet_images.rs,
   sheet_movies/caption.rs}`), then **chart captions** (Pages 3, Numbers 3:
   `{pages/editor/charts, numbers/editor/sheet_charts}/caption.rs`) on the
   lifted engine.

Patterns from this wave: the host appends absent fields in its call order,
so a shared rewrite that must be byte-identical applies components in that
order (`value_axis::ordered_writes`); the host keeps an unreferenced asset
only when the data metadata map lists it or a message info references it
(`data_owner::release`); each core's locality check is where a new
publication shape is admitted, gated on the scope's allowed fields. Cost
probe `chart_property_cost.rs` (`W4C_ROUTE`); receipt writer
`.coord/iwa-w4i/tools/receipt.py`; ledger flips with
`.coord/iwa-w4i/tools/flip_rows.py` and a JSON spec.
