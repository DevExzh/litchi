# 0426 — W4-M: Pages chart copy and removal, Numbers image-fill release, Keynote line ends on the shared rules, Pages first table

- Status: merge-ready on branch `uo2/iwa-w4m` (cut from trunk
  `feat/unified-ops-api` at `a077aeeb43`).
- Date: 2026-11-03.
- Plan: [0363](0363-iwa-monolith-exit-plan.md) (§0, §3.2 API shape, §4
  parity tiers, §8 D1); the predecessor's brief
  [0423](0423-iwa-w4j-last-leftovers.md) §5; [0418](0418-iwa-w4e-chart-lifecycle-and-shared-authors.md)
  §6; the image-fill release of [0415](0415-iwa-w4p-image-fill-release-and-next-scope.md);
  the shared `data_owner` of [0422](0422-iwa-w4i-value-axis-image-fills-captions.md);
  the from-scratch table graph of [0420](0420-iwa-w4g-shared-table-graph.md).
- Scope (coordinator, wave W4-M):
  1. Pages chart duplicate and remove (2 ledger rows);
  2. Numbers image-fill replace/reset releasing the old image;
  3. Keynote line endpoints on the shared `line_end` rules, Keynote's copy
     deleted, parity unchanged, cost within 5% of the current leaf;
  4. Pages table insertion into a document with no table.

  Keynote slide-table add on decks without a creation table stays out of
  scope: the host refuses it too.
- `performance_claim`: every new or changed route uses 75–92% fewer
  instructions than its host route; the Keynote endpoint routes after the
  move are within ±0.1% of the leaf before it (§5).

## 1. Pages chart duplicate and remove (2 rows)

| Commit | What |
|---|---|
| `293b85a242` *(shared)* | `litchi_iwa_style::chart_lifecycle::graph`: `references` lists a chart drawable's parent, comment, title, caption, active and own preset, mediator and its chart, legend, axis, series (sparse arrays through their entries) and reference-line styles in the host's `body_chart_graph` discovery order; `registered_style` is the host's stylesheet-registration rule (every chart style kind, and a series non-style whose style names a stylesheet); `remap` rewrites references in place at the known paths (`reference_paths` for the chart drawable, preset and mediator) and reports the references it saw, so an owner refuses a private reference outside them as the host's `remap_references` does |
| `e6d12ea881` *(host fix)* | `Draft::remove_body_chart` never pruned the removed attachment from the body storage header and refused itself ("remains referenced after deletion") on every chart, native or host-authored. It now prunes it as the media, shape and text-box removals do (record 0413). macOS-queued |
| `16cb665273` | Pages `drawable::lifecycle::Edit::{duplicate_chart(chart, anchor), remove_chart(chart)}` over `BodyChartSelector` (beside `add_chart`). The graph proof follows `body_chart_graph`: the drawable (parented to the body storage, no mediator), the caption slot (a stand-in, or the caption, its style when the drawable's member holds it, its storage and placement), the title stand-in, and when the chart owns its preset, the preset and every style only it references that its member holds or the stylesheet registers; the attachment last. A copy takes consecutive identifiers in that order, clones each object in its own member (chart drawable and preset remapped by `chart_lifecycle::graph`), moves the frame 12 pt and sets the attachment offsets from it and the left margin, joins the copied registered styles to the stylesheet and the preset to the theme with the host's cross-component references, anchors the copy with a U+FFFC and attachment entry, appends it to the z-order and registers the UUIDs (several components in one rewrite), watermark and data owners. A removal clears a direct comment first (`text_box_drawable_comment::clear_chart_comment`), deletes the U+FFFC and attachment entry, the z-order entry and the graph, unregisters the styles and preset (an emptied theme chart-preset list is dropped, as the host clears `extensions.chart`), and releases the cross-component references, data owners, UUIDs, orphaned data and identifier suffix. The media removal's orphan release becomes the shared `release_orphans` |
| `a1e2a4fd55` | parity, ledger and queue rows |

| Host route | Destination | Parity | Tier |
|---|---|---|---|
| `pages/editor/charts.rs duplicate_body_chart` | `litchi_pages::drawable::lifecycle::Edit::duplicate_chart` | `pages_body_chart_duplicate_parity` (the host-authored chart at two anchors, the native chart), `pages_body_chart_duplicate_native_parity` | B |
| `pages/editor/charts.rs remove_body_chart` | `litchi_pages::drawable::lifecycle::Edit::remove_chart` | `pages_body_chart_remove_parity` (both charts; the native document), `pages_body_chart_remove_commented_parity` | B |

Tier B: the host re-encodes the chart drawable, stylesheet, theme and
registry with Prost and draws random UUIDs; every other changed object
equals the host's on every known field and the copies take the host's
identifiers. `pages_body_chart_duplicate_then_remove_round_trips_reads`
removes two successive copies and gets the source's chart reads back;
`pages_body_chart_lifecycle_refusals` pins `NotFound` and the one-operation
rule. Both routes and the host fix are macOS-queued.

## 2. Numbers image-fill replace and reset release the old image

| Commit | What |
|---|---|
| `8c1e54f31b` | `litchi_iwa_style::shape::{stored_image, release_image_fills, StoredImage}`: the image fill a drawable's own style stores is captured before a fill change and, after the changes, released when its style no longer stores it: the style's message-info data reference, its registered owner through `data_owner::release`, and orphaned `DataInfo` and `Data/` member (the host's `remove_style_data_reference` and `remove_orphaned_image_asset`; the Pages route of record 0415 on the shared workspace). Numbers `sheet::drawable::style`: a fill set or reset over an image fill goes through `ReleasingFills`, and `set_image_fill` over one replaces it; `Refusal::ImageFill` now means only a second image fill staged in one edit. The deleted member joins the publication's deletions |
| `96475220d0` | `numbers_shape_image_fill_release_parity`: a solid fill, a reset and a new image over the host's image fill, Tier B, with the released `Data/` member checked on the leaf output |

No ledger row changes: the fill set, reset and image-fill routes were
already migrated; this closes their image-fill gap. macOS-queued.

## 3. Keynote line endpoints on the shared rules

| Commit | What |
|---|---|
| `48febe95c3` | `litchi_iwa_style::line_end::{endpoints, set_endpoints}` are generic over two new traits: `Staged` (the unique `ShapeInfo` payload, the shape's style, the exclusivity `Target`, and the `shape::Styles` reads) and `Staging` (plan publication). The shared `graph::Workspace` implements both, so the Numbers and Pages calls are unchanged. Keynote's slide-drawable shape styles implement them over Keynote's own staged workspace (`LineView` for reads, `LineStaging` for writes, publishing through the module's existing `apply`); Keynote's copies of the endpoint rules (`shape_line_ends`, `endpoint_of`, `inherited_line_ends`, `endpoints`, `set_endpoints`, about 140 lines) are deleted. Errors map to Keynote's typed errors exactly as before (an unknown line end and an orphan variation refuse; a missing read-back is `Verification`) |

The brief's alternative, a `graph::Workspace` over Keynote's components
with its `finish()` edits folded back, would rebuild the shared object index
and re-encode the registry for every endpoint change inside
`edit_slide_drawables`; making the shared rules generic over the staging
keeps Keynote's staging and costs nothing.

Parity is unchanged:

- `w4m_keynote_endpoints_fixed_digest` (leaf-built lines, so deterministic:
  a line created with arrows, set, rewrite, reset and a no-op set, with the
  reads) prints identical sizes, digests and reads on the source before the
  move (`cd3eb7181b`) and after it;
- every `keynote_slide_line_endpoints_*_parity` and
  `keynote_slide_line_*` case passes, as do the Numbers and Pages
  line-endpoint cases on the generic rules.

## 4. Pages table insertion into a document with no table

| Commit | What |
|---|---|
| `9aee95cccb` | `table::creation::Edit::insert` on a table-less document made from the built-in template (`Package::new`, `builder::Builder`; the host's scratch document) builds the host's from-scratch first table: `builder_table::install` (record 0420, on `litchi_numbers_wire::table_create::fresh`) now takes the anchor, and for an insertion appends the calculation-engine component after the registry's last component and the `CalculationEngine.iwa` member at the end of the package, where the host editor's `add_table` → `bootstrap_first_table_graph` → `add_component_registration` puts them (the builder keeps the host builder's layout). The name defaults to `Table 1`, the locale is the document root's (`en_US` when absent, as the host reads it), the UUID seed is the document identity's. Seeded cells go through the value writer as for any insertion. Any other table-less document still refuses `NoTemplate`, as the host does (`require_scratch_table_slots`) |
| `12225401ad` | `pages_body_table_insert_without_template_parity`: anchors 0, 6 and the end of the body text on a builder document, Tier B (table, owner, UID-map and object UUIDs by form, as for the builder), macOS-queued |

The leaf test `a_built_in_template_document_takes_its_first_table` covers
`Package::new` and a builder document at three anchors, reopening, the
exact inverse and a second insertion cloning the first table.
`invalid_requests_refuse_typed` still refuses `basic.pages` (native, no
table) with `NoTemplate`.

## 5. Performance

Callgrind, release parity binary with debug info, CPU 3, ABBA (two calls
of each route), inclusive instructions of the probe functions (open,
operation, commit, serialize). Raw captures in `.coord/iwa-w4m/cg{1,2,3}/`.

| Route | Probe | Host Ir | Leaf Ir | Leaf/host |
|---|---|---:|---:|---:|
| Pages chart duplicate | `pages_chart_lifecycle_cases::w4m_chart_lifecycle_cost_abba` (`duplicate`) | 2,527,294,110 | 211,752,846 | 0.084 |
| Pages chart remove | same (`remove`) | 1,915,631,038 | 192,546,368 | 0.101 |
| Numbers solid fill over an image fill (release) | `numbers_shape_style_cases::w4m_numbers_image_fill_release_cost_abba` | 303,635,523 | 94,422,861 | 0.311 |
| Pages first table (builder document, anchor 6, 3×4) | `pages_builder_cases::w4m_first_table_cost_abba` | 39,726,868 | 10,109,618 | 0.254 |

Keynote line endpoints, leaf before the move (`cd3eb7181b`, binary SHA-256
`88cd4452…2be7124`) against after it (`226a99c9ca`, `a783179b…f011d7`):

| Route (`w4m_keynote_endpoints_cost_abba`) | Host Ir | Leaf before | Leaf after | Δ leaf |
|---|---:|---:|---:|---:|
| set | 435,918,825 | 131,267,473 | 131,341,302 | +0.06% |
| rewrite | 313,658,778 | 120,638,243 | 120,626,799 | −0.01% |
| reset | 497,482,376 | 131,614,630 | 131,484,023 | −0.10% |

```sh
CARGO_PROFILE_RELEASE_DEBUG=1 cargo test --release -p litchi-iwa --test parity --all-features --no-run
taskset -c 3 valgrind --tool=callgrind target/release/deps/parity-<hash> --ignored --exact <probe> --test-threads=1
callgrind_annotate --inclusive=yes callgrind.out.<pid> | grep -E 'w4m_[a-z_]+_(host|leaf) '
```

## 6. Ledger

Two rows flip (`pages/editor/charts.rs duplicate_body_chart`,
`remove_body_chart`); Pages has 5 host-only rows left. Five W4-M rows join
`0363-mac-queue.md` (chart copy, chart removal, the host removal fix, the
Numbers image-fill release, the Pages first table). `python3
tools/iwa_exit_ledger.py --check` passes.

## 7. Left

| Item | Next |
|---|---|
| Keynote slide-table add on decks without a creation table | out of scope (the host refuses it too) |
| Pages image-fill release on the shared helper | Pages `body_drawable_style::release_images` keeps its own copy of the release transition (record 0415); it could move onto `shape::{stored_image, release_image_fills}`, but its orphan rule (any data metadata map keeps the asset) differs from `data_owner::release` (a map entry naming the data keeps it), so the move needs its own parity check |
| Pages first table on documents that are not the built-in template | the host refuses them as well (`require_scratch_table_slots`); no capability gap |
