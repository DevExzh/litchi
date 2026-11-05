# 0363 — Plan: dissolve `litchi-iwa` into the Pages, Numbers and Keynote leaf crates

- Status: plan (no code change). Branch `uo2/iwa-exit-plan`, cut from trunk
  `feat/unified-ops-api` at `25244be88f`.
- Date: 2026-10-26.
- Input: owner instruction of 2026-10-26 ("Refactor the iWork family
  according to ADR 0028 — move the Numbers, Pages, Keynote related logics to
  the leaf crates instead of a monolithic litchi-iwa"), the standing owner
  rules of [0350](0350-owner-api-debt-decision.md) (one consistent API, no
  capability loss, no deprecation or aliases, correctness, fidelity and
  performance preserved), contract [0351](0351-api-contract.md), ADR 0034 and
  `docs/API_GLOSSARY.md`, [ADR 0028](../adr/0028-iwa-monolith-exit.md) and its
  deletion gate, ADR 0002/0024, records [0352](0352-iwork-names.md),
  [0361](0361-w7-iwork-contract.md) and
  [0361a](0361a-w7-wire-and-structured-limits.md), the coordinator addition
  of 2026-10-26 (the unmerged Codex branches, §7).
- Evidence (small, committed): `results/0363/host-pub-fn-crosswalk.tsv`
  (every `pub fn` of `crates/litchi-iwa/src`, 2,228 rows, with format,
  family, destination crate, leaf parity and the leaf API cited) and
  `results/0363/crosswalk.py`, the classifier that produced it (§1.4).
- `performance_claim`: none; §5 lists the routes each package must measure.

## 0. Wave table

Every work package (WP) owns the paths listed and no others. Paths marked
"append-only" may receive only a separated `mod`/`pub use` hunk, the
rule the shared ops dispatch files already follow. "Donor" names the
unmerged Codex branch whose leaf code the WP ports instead of rewriting
(§7); a donor is never merged as is.

| WP | Wave | Scope | Owns | Donor (§7) | Size (host lines retired; donor lines ported) | Depends on | Performance routes (§5) | Deletion-gate items |
|---|---|---|---|---|---|---|---|---|
| **W0-A** substrate | 0 | admission, serialization reports and ZIP/IWA framing fixes the leaf families need; host physical raw access → `litchi-iwa-archive::unstable` | `crates/litchi-iwa-{archive,core,detect,package,cache}/**`, `crates/soapberry-zip/**`, `crates/litchi-iwa-common/src/wire.rs` | integration line `refactor/iwa-exit-current-20261023` substrate commits, `uo2/iwa-wire-budget-fix-20261024` | host `package.rs`, `bundle.rs`, `object_index.rs`, `package_metadata.rs`, `data_reference_registry.rs`, `snappy.rs`, `archive.rs`, `application_detection.rs` (~9.9k); ~15k ported | — | iWork open/save (0296, 0306) | 2 |
| **W0-B** codecs and neutral values | 0 | bounded Buffa codecs the leaf families need; neutral table/chart/shape/media values move out of `litchi-numbers` and the host into `litchi-iwa-common` (D4) | `crates/litchi-iwa-protos/**`, `crates/litchi-numbers-wire/**`, `crates/litchi-iwa-common/src/{table,chart,chart3d,shape,media,color,text}/**`; `crates/litchi-numbers/src/{cell,table}/**` value moves (re-export hunk only in `litchi-numbers/src/lib.rs`, append-only) | theme, shape-property, shape-creation, number-attachment and Numbers table-lifecycle codecs; `numbers-wire` formula graph/owner | host `protobuf.rs`, `registry.rs`, `theme.rs`, `numbers/{formula,formula_owner,function_map,table_uid_map}.rs` (~3.6k); ~12k ported | — | codec micro-routes only | 2 |
| **W0-C** parity harness and crosswalk gate | 0 | differential harness (host vs leaf, §4) inside the host's own tests; `tools/check_iwa_exit_crosswalk.py` making the crosswalk a checked ledger whose rows flip as WPs land | `crates/litchi-iwa/tests/exit_parity/**`, `crates/litchi-iwa/tests/exit_parity.rs`, `crates/litchi-iwa/Cargo.toml` (dev-deps only), `tools/check_iwa_exit_crosswalk.py`, `tools/test_check_iwa_exit_crosswalk.py`, `docs/unified-ops/results/0363/**` | — | new ~3k | — | harness overhead not measured | 2, 3 |
| **W1-A** Numbers tables and cells | 1 | table duplicate/resize, sheet duplicate, cell comment fallbacks, conditional highlights, cell formats (generic `DataFormat`, numeral system), cell rich text/paragraph/list, cell fill/border/layout, hidden axes, formula AST construction, pivot categories, Builder `table_dimensions`/`language`/`locale` | `crates/litchi-numbers/**` except `src/package/sheet_{drawable,shape,image,movie,audio,chart}*` and `src/{shape,sheet/chart}*`; `crates/litchi-numbers-wire/**` (the shared table engine W1-B and W1-C consume) | `refactor/iwa-table-lifecycle-native-tests-20261023` + `test/formula-owner-budget-20261023`, `refactor/numbers-conditional-highlight-admission-review-20261024`, integration-line `sheet_duplication.rs` | host `numbers/editor/{semantic/table.rs, semantic/workbook.rs, model.rs, storage.rs, cell_*, formula_*, conditional_highlight*, table_*, row_insert*, column_insert*, sheet_delete.rs, sheet_duplicate*, selectors.rs, tests.rs}`, `numbers/creation.rs` (~62k) | W0-A, W0-B, W0-C | cell rich-text set 1/1k, conditional-highlight set, table/sheet duplicate, Builder build | 3, 4 |
| **W1-B** Pages body, sections and tables | 1 | body media (image/movie/audio lifecycle and properties), body shapes, table comments, number attachments, bookmarks, body date-time fields, body/drawable text storages, table cell formats/rich text/appearance/conditional highlights/formulas/topology/duplicate, Builder `language`/`locale`/`body_table` | `crates/litchi-pages/**` except `src/package/{text_box,body_chart_lifecycle,body_chart_style}*` | `refactor/pages-media-lifecycle-admission-20261024`, `repair/pages-body-shape-a408-20261024` (+ `refactor/pages-body-shape-review-20261024` history), `feat/pages-table-comment-ownership-20261024`, `refactor/iwa-attachment-reader-admission-20261023`, `pages-construction-boundary-20261024`, integration-line `drawable_comments.rs` and `package_builder.rs` | host `pages/editor/{images*, movies*, audio*, body_shapes*, tables/**, number_attachments.rs, section_content.rs, footnotes.rs, media_playback.rs, image_adjustments.rs, tests.rs}`, `pages/creation*` (~33k) | W0-A, W0-B, W0-C | shape add, image add, table cell style set 1/1k, Builder build | 3, 4 |
| **W1-C** Keynote slides, drawables and media | 1 | slide-table lifecycle/geometry, chart lifecycle, image lifecycle, shape and text-box lifecycle, layout reassignment, builds and effects, table cell rich text/formats/comments/highlights/formulas/hidden axes, public table topology, Builder `language`/`locale` | `crates/litchi-keynote/**` except `src/package/{slide_text_box_text,slide_chart_style}*` | `feat/keynote-table-lifecycle-focused-20261024`, `repair/keynote-chart-deploy-20261024`, `refactor/keynote-shape-text-repair-20261024-r5`, integration-line image lifecycle, layout reassignment and `package_builder.rs` | host `keynote/editor/{slide_tables*, slide_create*, slide_layout_*, slide_images*, slide_shapes*, builds.rs, slide_graph.rs, keynote_object_catalog.rs, placeholder_ownership.rs, tests*}`, `keynote/creation*` (~37k) | W0-A, W0-B, W0-C | chart create, table create, shape add, build add, Builder build | 3, 4 |
| **W1-D** shared text engine | 1 | TSWP storage transformations (character and paragraph properties, lists, drop caps, languages, hyperlinks, highlights, comments, bookmarks, date-time fields, number attachments, named paragraph styles) as storage-scoped engines over semantic values | `crates/litchi-iwa-text/**`, `crates/litchi-iwa-text-wire/**` | `uo2/iwa-storage-decode-admission-20261023` (superseded; read only) | host `text/**` (~34k) | W0-A, W0-B | storage-level property set 1/1k runs | 3 |
| **W1-E** shared chart, shape and media engine | 1 | chart style/axis/legend/series/3D/pie/radar codecs, chart lifecycle graph clone, shape fill/stroke/shadow/effects/path/line-end/text-layout codecs, title/caption construction, media asset discovery and replacement | `crates/litchi-iwa-common/src/{chart,chart3d,shape,media}/**` (after W0-B), new `crates/litchi-iwa-protos/src/{chart_style,shape_style,caption,media_asset}_codec.rs` (append-only `lib.rs` hunk) | `repair/keynote-chart-deploy-20261024` shared parts | host `charts/**`, `shapes/**`, `image_caption.rs`, `media/**`, `drawable_order.rs`, `table_hidden_axes.rs` (~49k) | W0-B | chart property set, shape property set | 3 |
| **W2-A** Numbers sheet drawables | 2 | sheet text boxes (lifecycle, properties, rich text via W1-D), shapes, images, movies, audio, charts (lifecycle and style via W1-E), drawable order, drawable comments | `crates/litchi-numbers/src/package/sheet_{drawable,shape,image,movie,audio,chart,text_box}*`, `src/{shape,sheet/chart}*`, `crates/litchi-numbers/examples/**` | `feat/numbers-drawable-comments-focused-rereview-20261024` | host `numbers/editor/{sheet_*, semantic/{text_box,drawables,media}.rs, text_box_*, date_time_fields.rs, named_paragraph_styles.rs, drawable_order.rs, cell_paragraph_*, stroke_layers.rs}` (~38k) | W1-A, W1-D, W1-E | text-box style set 1/1k, chart create/style, image add | 3, 4 |
| **W2-B** Pages text boxes and charts | 2 | text boxes and drawable rich text, named paragraph styles, chart lifecycle and style | `crates/litchi-pages/src/package/{text_box,body_chart_lifecycle,body_chart_style}*`, `crates/litchi-pages/examples/**` | — | host `pages/editor.rs`, `pages/editor/{charts*, text_box_create.rs, date_time_fields.rs, types.rs}` (~21k) | W1-B, W1-D, W1-E | text-box style set 1/1k, chart create/style | 3, 4 |
| **W2-C** Keynote text boxes and charts | 2 | slide text-box rich text, named paragraph styles, date-time fields, chart style families | `crates/litchi-keynote/src/package/{slide_text_box_text,slide_chart_style}*`, `crates/litchi-keynote/examples/**` | — | host `keynote/editor.rs`, `keynote/editor/{slide_charts*, text_box_create.rs, named_paragraph_styles.rs, date_time_fields.rs, drawable_order.rs, operation_cache_tests.rs}` (~19k) | W1-C, W1-D, W1-E | text-box style set 1/1k, chart style set | 3, 4 |
| **W2-D** tests, examples, fixtures and docs relocation | 2 | the 18 host integration targets and remaining host unit tests move to leaf owners or are retired as duplicates; the 206 host examples are rewritten as leaf examples or retired with their capability's leaf example; fixture provenance moves off the host | `crates/litchi-iwa/tests/**` (except `exit_parity*`), `crates/litchi-iwa/examples/**`, `test-data/iwork/**/README.md`, `crates/litchi-{pages,keynote,numbers}/assets/README.md`, `docs/IWORK_*_AUDIT.md` | the shared base commits `b257dd7ba2..5818f7613f` (host example retirement, oracle moves) | host tests 10.6k, examples 15.8k | W1-* (a test moves when its family has landed) | none | 2 |
| **W3** deletion and ledger | 3 | delete `crates/litchi-iwa`; empty `migration_hosts`/`migration_debt`; prune the host audits of the boundary tool and its tests; delete the `iwa` glossary profile; update `non_iwork_gate`, `raw_api_census_matrix.json`, `validate_api_disposition.py`, `check_iwork_leaf_public_api.py`; ADR 0028 status, ADR 0024 topology amendment | `crates/litchi-iwa/**` (delete), `tools/{crate_boundaries.json, check_crate_boundaries.py, test_check_crate_boundaries.py, api_glossary.json, non_iwork_gate.py, test_non_iwork_gate.py, raw_api_census_matrix.json, validate_api_disposition.py, check_iwork_leaf_public_api.py, test_check_iwork_leaf_public_api.py, api_debt_inventory.py}`, `docs/adr/{0028,0024}*.md`, `docs/adr/README.md` | — | host remainder; tool pruning ~2k of 1,814 host references | all | workspace build only | 1, 5, 6 |

Wave 0 runs three agents; waves 1 and 2 run up to five (W1-A…E, then
W2-A…D); wave 3 runs one. Within a wave no two WPs own the same path: the
three format WPs own one leaf crate each, W1-D/W1-E own the shared crates,
W2-D owns the host's tests and examples. A leaf crate's `lib.rs`,
`package.rs` and `Cargo.toml` belong to its format WP; W0-B and W1-E add only
separated hunks. Each wave ends with a coordinator merge and a full gate
before the next starts.

## 1. Inventory of `crates/litchi-iwa`

### 1.1 Size and public surface

302k source lines, 10.6k test lines (18 integration targets, 83 tests),
15.8k example lines (206 examples), 1,697 unit tests. Public modules
(`lib.rs`): `application`, `charts`, `comments`, `document`,
`drawable_order`, `keynote`, `media`, `numbers`, `pages`, `shapes`, `text`;
root re-exports `Document`, `Package`/`Edit`/`Commit`, `Patch`,
`Diagnostics`, `LimitKind`, `DrawableLayerMove`, `DrawableTitleCaption`,
the media types and the text values. `archive`, `bundle`, `object_index`,
`package`, `package_metadata`, `protobuf`, `registry` (test-only), `snappy`,
`theme`, `wire`, `identity`, `image_caption`, `table_hidden_axes`,
`data_reference_registry` and `application_detection` are private. The
`pub` scan (`grep -rnoE '^\s*pub fn [a-z0-9_]+'`) finds 2,228 public
functions; the three `Draft` editors carry 1,900 of them and take native
object identifiers (`table_id: u64`, `drawable_object_id: u64`,
`slide_index: usize`) — the "Native-ID host API" of record 0361 §5.

### 1.2 Modules

Leaf parity: **Y** the leaf crate (or a focused crate) implements the
capability at parity today; **P** partially (named leaf API covers part,
the rest is host-only); **B** implemented on an unmerged Codex branch only
(§7); **N** host-only. "Lines" are source lines including in-file tests.

| Module | Lines | Public surface | What it does | Internal deps | Leaf / focused parity |
|---|---:|---|---|---|---|
| `numbers::creation` | 2.0k | `numbers::Builder` (`new, sheet_name, table_name, table_dimensions, language, locale, build`) | source-free `.numbers` authoring | `numbers::editor`, `package`, protos | **P** `litchi_numbers::create::Builder` (`identity, sheet_name, table_name, build`); no `table_dimensions`, `language`, `locale` |
| `numbers::editor` (semantic/table, model, storage, cell_*) | 47k | `numbers::Draft` table-cell methods (109 in `semantic/table.rs`) and 50 value types (`TableCellText*`, `TableCellParagraph*`, `TableSort*`, …) | per-cell rich text, paragraph/list properties, fill, borders, layout, data formats, numeral system, merges, cell comments | `numbers-wire` (BNC), `litchi_numbers` values, protos, `text` | merges **Y** (`Package::edit_table_merges`, `MergeReader`); character style **P** (`edit_table_cell_text_style`); 11 dedicated formats **P** (`edit_table_cell_{number,currency,percentage,scientific,fraction,date_time,duration,text,custom,control,pop_up_menu}_format`), generic `DataFormat` and numeral system **N**; paragraph/list, per-cell fill/border/layout **N**; cell comments **P** (`edit_table_cell_comment[_replies]`; storage-creation and identity fallbacks **N**) |
| `numbers::editor` (workbook, sheet_delete/duplicate, table_*) | 12k | `open, from_bytes, sheets, tables, add_empty_sheet, remove_sheet, duplicate_sheet, add_empty_table, duplicate_table, remove_table, resize_table, insert/remove_table_row/column, set_formula*` | workbook structure | as above | sheet add/remove **Y** (`edit_sheet_lifecycle`), duplicate **B**; table add/remove **Y** (`edit_table_lifecycle`), duplicate/resize **B**; row/column **Y** (`edit_table_topology`); order **Y** (`edit_sheet_order`, `edit_table_order`); formulas **P** (`edit_table_cells` `set_formula[_cached]`; AST construction `numbers::FormulaExpression` **N**) |
| `numbers::editor` (table_sort, table_hidden_axes, conditional_highlight, formula_cache/clone/dependency_shift) | 15k | `table_sort_order…apply_table_sort_order_to_rows`, `table_hidden_axes`, `cell_conditional_highlight*` | sort, hidden rows/columns, conditional highlighting, formula maintenance | `numbers-wire` | sort **Y** (host delegates to `edit_table_sort_order`/`apply_table_sort_order`); hidden axes **N**; conditional highlights **B**; formula maintenance is internal to the families above |
| `numbers::editor` (sheet_shapes, sheet_images, sheet_movies, sheet_audio, sheet_charts, text_box_*, named_paragraph_styles, date_time_fields, drawable_order) | 30k | 40 shape, 16 image, 19 movie, 11 audio, 186 chart, 139 text-box, 8 style, 5 field, 3 order methods | sheet drawables | `charts`, `shapes`, `text`, `media`, `image_caption` | chart data/arrangement **P** (`edit_sheet_chart_{data,arrangement}`), image adjustments **P** (`edit_sheet_image_adjustments`); everything else **N**; drawable comments **B** |
| `numbers::formula` | 1.6k | `FormulaExpression`, `FormulaCellReference`, … | typed formula AST to native archive | protos | **N** (vocabulary is already `litchi_numbers::formula`; construction is host-only) |
| `pages::creation` | 2.3k | `pages::Builder` (`new, body_text, language, locale, body_table, build`) | source-free `.pages` authoring | `pages::editor`, `numbers::editor` | **P** `litchi_pages::builder::Builder` (`identity, body_text`), `Package::new`; no `language`, `locale`, `body_table`; integration line adds `package_builder.rs` (**B**) |
| `pages::editor` (editor.rs) | 5.9k | `pages::Draft` body, section and text-box methods (179) | body storage, sections, text-box rich text, drawable comments | `text`, `comments`, `media` | sections **Y** (`edit_section_lifecycle`); body/section text **P** (`edit_body_text`, `edit_section_text`, `edit_section_text_style`, `edit_section_paragraphs`, `edit_header_footer_text`); drawable text storages and text boxes **N**; drawable comments **B** |
| `pages::editor::tables` | 8.4k | 157 table-cell methods + topology, formula, comments, sort, conditional highlight | attached body tables (reuses the Numbers table engine) | `numbers::editor`, `litchi_numbers` values | values **Y** (`edit_body_table_values`), merges **Y**, sort **Y**, hidden axes **Y**, creation/deletion **Y** (`edit_body_table_{creation,deletion}`), headers/lock/name/title/appearance/dimension **Y**; cell formats, rich text, fill/border/layout, conditional highlights, formulas, row/column topology, duplicate **N**; table comments **B** |
| `pages::editor` (images, movies, audio, media_playback, image_adjustments) | 6.5k | 16 image, 19 movie, 11 audio methods | body media | `media`, `image_caption` | image adjustments **Y** (`edit_body_image_adjustments`); lifecycle and properties **B** |
| `pages::editor::charts` | 14k | 184 chart methods | body charts | `charts` | data and arrangement **P** (`edit_body_chart_{data,arrangement}`); lifecycle and style **N** |
| `pages::editor` (body_shapes, footnotes, number_attachments, section_content, text_box_create, date_time_fields) | 6.4k | 45 shape, 12 attachment, 12 section-content methods, footnotes | body shapes, footnotes, attachments, bookmarks, date-time fields | `shapes`, `text` | footnotes **Y** (`edit_body_footnote`, `edit_body_footnote_text`); shapes **B**; number attachments **B**; bookmarks and date-time fields **N** |
| `keynote::creation` | 2.4k | `keynote::Builder` (`new, title, subtitle, presenter_notes, language, locale, slide_size, build`) | source-free `.key` authoring | `keynote::editor` | **P** `litchi_keynote::Builder` (`identity, title, subtitle, notes, slide_size`); no `language`, `locale` |
| `keynote::editor` (editor.rs, builds.rs) | 7.6k | `keynote::Draft` text-box methods (203), build/effect values (52) | slide text boxes, builds, media assets | `text`, `media` | slide text **P** (`edit_slide_{text,body,title,notes}`); build order **P** (`edit_slide_build_order` move only); add/set/remove build and effects **N**; text-box rich text **N** |
| `keynote::editor::slide_tables` | 9.8k | 155 table methods | slide tables | `numbers::editor` | merges, sort (incl. physical), headers, lock, name, title, appearance, dimension, number format **Y** (`edit_slide_table_*`); cells **Y** (`edit_slide_drawables` `set_table_cell`); lifecycle **P** (`add_table`/`remove`; duplicate, resize, geometry **B**); rich text, other formats, fill/border/layout, comments, highlights, formulas, hidden axes **N**; topology **P** (crate-internal) |
| `keynote::editor::slide_charts` | 12.7k | 168 chart methods | slide charts | `charts` | data, arrangement, title, caption, legend, axis title, value axis **Y** (`edit_slide_chart_*`); lifecycle **B**; style **N** |
| `keynote::editor` (slide_create, slide_layout_*, slide_images, slide_shapes, slide_graph, object catalog, placeholders, drawable_order) | 9.9k | `add_slide, insert_slide, duplicate_slide, set_slide_layout`, 16 image, 45 shape, 3 order methods | slide lifecycle and drawables | `media`, `shapes` | slide add/duplicate/remove **Y** (`edit_slide_creation`, `edit_slide_deletion`), layout **B**; drawables **P** (`edit_slide_drawables`: `add_{image,shape,table,text_box}`, `remove`, `set_geometry`, `move_to`, `set_text[_style]`); image lifecycle **B**; shape lifecycle **B**; shape style **N**; movies/audio **Y** (`edit_slide_movie_*`, `edit_slide_audio_*`, `edit_slide_media_*`); drawable comments **Y** (`edit_slide_drawable_comment`) |
| `charts` | 33.8k | 49 chart value types (`Chart3dDepth`, `ChartLegendFont`, `ChartSeriesTrendline`, …), `reference_line` | per-property chart archive codecs shared by the three editors | protos (generated Prost) | values **P** (`litchi_iwa_common::chart::{Axis, Gridlines, Steps, Scale, NumberFormat, gaps, pie, reference_line, …}`); legend, font, fill, stroke, shadow, symbol, trendline, 3D, radar, donut values and every style codec **N** |
| `text` | 33.7k | `text::Draft`, `TextExtractor`, `TextStorageInfo`, `AppliedParagraphStyle`, `TextDateTimeField` | TSWP storage transformations (alignment 9.8k, lists 7.2k, drop caps, languages, hyperlinks, highlights, comments, bookmarks, date-time fields, number attachments, paragraph-style catalog) | protos, `litchi-iwa-text`, `text-wire` | semantic values **Y** (`litchi_iwa_text::{ParagraphList*, DropCap, TextHyperlink, TextHighlight, TextComment, Bookmark, TextNumberAttachment, NamedParagraphStyle, …}`); character-property rewrite **P** (`litchi_iwa_text_wire::char_style`); paragraph and list rewrites, attachments, fields **N** |
| `shapes` | 7.9k | `DrawableGeometry`, `DrawableProperties`, `ShapeFill`, `ShapeImageFill`, `LineSegment`, `ShapePathKind`, `DrawableTitleCaption` | shape fill/stroke/shadow/effects/path/line-end/text-layout codecs | protos | values **P** (`litchi_iwa_common::shape::{Stroke, Shadow, Effects, Gradient, Endpoints, Preset, …}`); codecs **N** (Pages shape codec **B**) |
| `media` | 3.6k | `MediaManager`, `MediaAsset`, `MediaAssetId`, `EmbeddedMediaAsset`, `MediaStats`, `media::{Draft, OpenOptions}` | `Data/*` discovery, extraction, replacement | `package`, `bundle` | Keynote slide media **Y** (`edit_slide_media_data`); generic asset listing/extract/replace **N** (Pages/Numbers media **B**, `litchi-iwa-media` on the Pages media branch) |
| `comments`, `comments.rs` | 4.8k | `comments::Draft` | drawable comments shared by Pages and Numbers | `text`, protos | Keynote **Y**; Pages, Numbers **B** |
| `document` | 0.5k | `Document::{open, from_bytes, text, charts, application, validate, media_stats, extract_media*, stats, message_type_summary}` | read-only whole-package view | all | text **P** (`litchi_pages::Package::text`, Keynote/Numbers semantic text); chart metadata **Y** (leaf `*_chart_metadata`); message-type statistics **N** |
| `package` | 2.7k | `Package` (`open, from_bytes, from_shared_bytes, from_read_at, edit, apply, entry_names, iwa_entry_names, entry, archive, validate, to_bytes, write_to, save`), `Edit`, `Commit` | mutable entry-order-preserving ZIP package with archive re-encoding | `litchi-iwa-archive`, `core` | physical **P** (`litchi_iwa_archive::{package, publication, unstable}`); raw entry/archive mutation **N** as a supported API (D5) |
| `bundle`, `object_index`, `package_metadata`, `data_reference_registry` | 7.3k | private | bundle layout, object locations, object/UUID/data registries | archive, index | **Y** `litchi_iwa_archive::bundle`, `litchi_iwa_index::ObjectIndex`; registry rewrites **P** (leaf families carry private copies) |
| `protobuf`, `registry`, `theme`, `wire`, `snappy`, `archive`, `identity`, `application*`, `drawable_order`, `image_caption`, `table_hidden_axes` | 5.9k | `DrawableLayerMove`, `Application`; rest private | Prost helpers, type registry, theme extensions, re-exports, fresh identities, detection, shared drawable/caption/axis helpers | protos, archive | detection **Y** (`litchi_iwa_detect`); identities **Y** (leaf `Identity`); snappy/archive/wire **Y** (core/archive/common); theme **B** (theme codecs); drawable order **P**; captions **P** (Keynote/Numbers/Pages chart and movie caption families); Prost helpers disappear |

### 1.3 Per-format totals (crosswalk)

| Format | Public fns | Y | P | B | N |
|---|---:|---:|---:|---:|---:|
| Numbers | 594 | 19 | 145 | 11 | 419 |
| Pages | 667 | 27 | 53 | 101 | 486 |
| Keynote | 639 | 18 | 217 | 8 | 396 |
| shared modules | 328 | 59 | 0 | 0 | 269 |
| **total** | **2,228** | **123** | **415** | **120** | **1,570** |

The counts are per host function, not per capability: one host family of
40 `set_x`/`reset_x` methods (for example the text-box character
properties) becomes one leaf `Edit` verb taking a property value, so the
leaf surface will be far smaller than 2,228.

### 1.4 The crosswalk file

`results/0363/host-pub-fn-crosswalk.tsv` has one row per host `pub fn`:
`path:line`, name, format, family, destination crate, parity (`Y/P/B/N`)
and evidence (the leaf API or branch). Regenerate with:

```sh
cd crates/litchi-iwa/src
for m in numbers keynote pages charts text shapes media comments.rs comments \
         document.rs drawable_order.rs application.rs lib.rs package.rs; do
  grep -rnoE '^\s*pub fn [a-z0-9_]+' $m | sed -E 's/:\s*pub fn /\t/'
done > /tmp/pubfns.tsv
python3 docs/unified-ops/results/0363/crosswalk.py /tmp/pubfns.tsv
```

The family is assigned by name pattern and is a planning aid: each WP's
first step reviews its rows, corrects the family and parity per verb, and
fills a `disposition` column (moved to `leaf::path`, merged into verb,
retired as duplicate of …). W0-C turns that column into a checked gate.

## 2. Who uses `litchi-iwa`

### 2.1 Manifests and code

- **No manifest depends on it** (deletion gate 1 already holds): no
  `Cargo.toml` in the workspace, the fuzz crates (`crates/*/fuzz`),
  `tools/perf-baseline`, `tools/native-resave` or any `Cargo.lock` names the
  package; the root `Cargo.toml` has no `[workspace.dependencies]` entry; the
  package is a member only through `members = ["crates/*"]`. No CI workflow,
  hook or script runs `-p litchi-iwa`.
- **Rust code outside the crate:** none. Three files mention it in comments
  only: `crates/litchi-pages/src/builder.rs` (template provenance),
  `crates/litchi-numbers/src/create.rs` (writer provenance),
  `crates/litchi-keynote/tests/slide_drawable_comments_host_parity.rs`
  (describes the graph shapes the old `keynote::Draft` accepted; it does not
  link the host).
- **The crate's own users:** its 206 examples, 18 integration targets and
  1,697 unit tests.
- **Fixture provenance that names the host:** the embedded Pages template
  (`litchi-pages/src/builder.rs`, record 0299), the Keynote blank
  presentation (`crates/litchi-keynote/assets/README.md`: generated by
  `litchi_iwa::keynote::Builder` at `802995d7b7`), the synthetic Numbers
  workbook (`test-data/iwork/synthetic/numbers/README.md`), and
  `test-data/iwork/README.md` (lines ~694–699, ~1230). These are frozen
  bytes; they keep their recorded generator commit. Deleting the host does
  not change them, but their "regenerate" recipes must point at a
  pre-deletion commit or at the leaf `Builder` once it reproduces the bytes
  (W2-D).
- **Tools that name the host:** `tools/check_crate_boundaries.py` (1,099
  references) and its test (715), `tools/api_glossary.json` (`iwa` profile,
  400), `tools/non_iwork_gate.py` (40) and its test (42),
  `tools/raw_api_census_matrix.json` (42), `tools/validate_api_disposition.py`
  (10), `tools/check_iwork_leaf_public_api.py` (3, as a forbidden dependency)
  and its test, `tools/api_debt_inventory.py` (1). W3 prunes them.
- **Docs:** `crates/litchi-iwa/README.md`, the three leaf
  `docs/FEATURE_MATRIX.md` "legacy-host delta" sections,
  `docs/IWORK_FEATURE_MATRIX_AUDIT.md`, `docs/IWORK_PROGRESS_AUDIT.md`,
  ADR 0024 and ADR 0028. Dated records are history and are not rewritten.

### 2.2 The 11 migration-debt edges (`tools/crate_boundaries.json`)

| Order | Edge | Reason (ledger) | Exit condition (ledger) | Closed by |
|---:|---|---|---|---|
| 1 | `litchi-iwa → litchi-core` | unified document, media and metadata adapters over core primitives | move each adapter into its format owner or the root facade, then delete | W1-*, W1-E (media), W3 |
| 2 | `→ litchi-iwa-archive` | legacy mutation paths coordinate package ingress, preservation, reassembly | move all package transactions into Pages/Numbers/Keynote, delete | W0-A, W1-*, W2-*, W3 |
| 4 | `→ litchi-iwa-common` | unmigrated editors consume neutral values and wire helpers | move every editor adapter to its owner, delete | W1-*, W2-*, W3 |
| 8 | `→ litchi-iwa-index` | legacy object index bridges raw archives to editors | adopt the focused index inside the owners, delete | W0-A, W1-*, W3 |
| 10 | `→ litchi-iwa-protos` | editors decode generated protobuf payloads directly | replace with private bounded format codecs, delete | W0-B, W1-E, W1-D, W3 |
| 12 | `→ litchi-iwa-text` | text editors adapt native storage graphs to semantic text | move application text adapters into the owners, delete | W1-D, W2-*, W3 |
| 13 | `→ litchi-iwa-text-wire` | host invokes the shared text-storage wire adapter | route every call through the format packages, delete | W1-D, W2-*, W3 |
| 14 | `→ litchi-keynote` | Keynote native mutation remains in the host | complete Keynote package editing, move examples and tests, delete | W1-C, W2-C, W2-D, W3 |
| 15 | `→ litchi-numbers` | substantial Numbers native editing remains in the host | complete Numbers package editing, move examples and tests, delete | W1-A, W2-A, W2-D, W3 |
| 16 | `→ litchi-numbers-wire` | host table editors (all three formats) use the BNC codec | move all table mutation adapters into the owners, delete | W1-A/B/C, W3 |
| 17 | `→ litchi-pages` | Pages native mutation remains in the host | complete Pages package editing, move examples and tests, delete | W1-B, W2-B, W2-D, W3 |

The edges are removed together in W3: the host keeps compiling against all
of them until it is deleted, and the parity harness (§4) needs the host to
link the leaf crates. `parse_policy` accepts an empty `migration_hosts` only
when `migration_debt` is empty too, so W3 empties both in one commit.

## 3. Ownership mapping and API shape

### 3.1 Destination per module (ADR 0028 "Residual work moves by ownership")

| Host area | Destination | Visibility |
|---|---|---|
| ZIP preservation, raw entry listing/reading, message-type statistics (`Package::{entry_names, entry, archive, iwa_entry_names}`, `Document::{stats, message_type_summary}`) | `litchi-iwa-archive::unstable` (D5) | explicitly unstable physical vocabulary, not re-exported by any format facade |
| raw entry mutation (`insert_entry`, `remove_entry`, `replace_archive`, `update_archive`; already crate-private) | deleted with the host; leaf families publish through `litchi-iwa-archive::publication` | — |
| IWA framing, Snappy, archive metadata, identity remap | `litchi-iwa-core` | low-level, unstable |
| object index, references, `package_metadata`/`data_reference_registry` rewrites | `litchi-iwa-index`, `litchi-iwa-graph`; registry rewrites as one private shared helper in `litchi-iwa-archive` instead of per-family copies | private to the owners |
| detection, application classification | `litchi-iwa-detect` | existing |
| TSWP storage transformations (`text/**`) | `litchi-iwa-text-wire` (wire engines) over `litchi-iwa-text` values | `text-wire` low-level; the values are the supported vocabulary the leaf families take |
| chart, shape, caption, media codecs (`charts/**`, `shapes/**`, `image_caption.rs`, `media/codec.rs`, `theme.rs`) | bounded Buffa codecs in `litchi-iwa-protos`; values in `litchi-iwa-common::{chart, chart3d, shape, media}` | codecs private to the owners (protos is never re-exported); values public |
| shared TST/BNC table engine (`numbers/editor/{model,storage,cell_*,formula_*}`, used by all three formats) | `litchi-numbers-wire` (private to the owners, ADR 0028) | low-level |
| neutral table vocabulary the Pages and Keynote tables use today from `litchi_numbers` (data formats, topology operations, `Region`, cell `Value`) | `litchi-iwa-common::table` with `litchi-numbers`, `litchi-pages` and `litchi-keynote` re-exporting it at their existing paths (D4) | public |
| Numbers / Pages / Keynote topology, selectors, transactions, native mutations | `litchi-numbers`, `litchi-pages`, `litchi-keynote` | supported facade |
| `Document` whole-package text/charts/media | leaf `Package` reads; cross-format aggregation stays `litchi-iwa-structured` and the root `litchi::iwork` coordinator | supported |
| builders | leaf `Builder`s | supported |
| `numbers::editor::topology_oracle::__edit_table_topology_unverified` | test-only helper of the leaf topology tests | not public |

### 3.2 The shape every migrated capability takes (0351)

- **Owner.** The leaf `Package` is immutable; a capability is a scoped family
  `x` in a focused module `x::{Edit, Commit, Patch, Diagnostics, Error,
  LimitKind}`: `Package::x(&self, selector)` reads,
  `Package::edit_x(&self[, selector]) -> x::Edit` stages,
  `edit_x_with_options(.., x::EditOptions)` configures,
  `x::Edit::commit(self) -> Result<x::Commit>` publishes a new `Package`
  with `Patch` and `Diagnostics`, `Package::apply_x(&self, &x::Patch)`
  replays, `patch.inverse()` undoes; an exact no-op returns the source
  (`diagnostics().changed() == false`, `patch.is_noop()`).
- **Property families collapse.** A host run of `x()`/`set_x()`/`reset_x()`
  methods becomes one family with a property-valued verb, as
  `litchi_pages::Package::edit_section_text_style` and
  `litchi_numbers::Package::edit_table_cell_text_style` already do
  (`Edit::set(span | At, CharacterProperties)`). Planned families
  (names per ADR 0034; each WP registers them in the glossary before use):
  - text boxes: `edit_{sheet,body,slide}_text_box_text(selector)` with
    `set_character(span, CharacterProperties)`,
    `set_paragraph(paragraph range, ParagraphProperties)`,
    `set_list(..)`, `reset_*`, and hyperlink/highlight/comment/language/
    drop-cap/date-time-field/number-attachment verbs;
  - table cells: `edit_*_table_cell_text(selector, table)` (character and
    paragraph), `edit_*_table_cell_format` (one family taking a format enum
    replaces the 14 per-format host families; Numbers keeps its existing
    dedicated families and adds the missing kinds), `edit_*_table_cell_appearance`
    (fill, border, layout);
  - charts: `edit_*_chart_lifecycle` (create, duplicate, remove, kind,
    direction, geometry) and `edit_*_chart_style(selector)` with
    property verbs for axis, legend, series, 3D, pie and radar values;
  - shapes, images, movies, audio: `edit_*_{shape,image,movie,audio}_lifecycle`
    and `_properties`, following the existing Keynote movie/audio families.
- **Errors.** `Error::Refused(Refusal)` with a typed module refusal and
  `kind()`, `PatchConflict`, `SourceChanged`, `Limit { kind, observed,
  maximum }`, `InvalidLimit`, internal variant; no `String` flattening (the
  host's `Error::{InvalidFormat, ParseError}(String)` routes go away).
- **Limits** through `OpenOptions` and `EditOptions`; no `*_with_limits`
  (the donors use 14–77 `_with_limits` names per file and `LimitExceeded`
  variants; they are renamed when ported, §7).
- **Selector-first.** Every target is a semantic selector:
  `SheetSelector`/`TableSelector`/`ChartSelector`/`ImageSelector` (Numbers,
  existing), `SectionSelector`/`BodyTableSelector`/`BodyChartSelector`/
  `ImageSelector` (Pages, existing), `SlideSelector`/`DrawableSelector`/
  `TableSelector`/`ChartSelector`/`MovieSelector` (Keynote, existing). New:
  Numbers `DrawableSelector` (sheet + kind + position or name), Pages
  `DrawableSelector` (body or section + kind + position), and text-box
  selectors in all three. No native object identifier, UUID or protobuf
  value appears in a supported facade signature; the host's `*_object_id:
  u64` arguments are resolved by the families' private ownership proofs.
- **Raw-ID capabilities** (D5): the host's documented "Native-ID host API"
  (record 0361 §5.2: drawable comments for Pages/Numbers, cell-comment
  fallbacks) is re-expressed through selectors, which the donors already do
  for drawable and table comments. Physical inspection by identifier stays
  available through `litchi-iwa-index::ObjectIndex` and
  `litchi-iwa-archive::unstable`, whose vocabulary is explicitly unstable
  and never glob-re-exported. No raw-ID mutation API survives.

## 4. Parity strategy without native Apple applications

This Linux host cannot run Pages, Numbers or Keynote. The host
implementation is therefore the oracle, and it stays linked until W3.

### 4.1 The differential harness (W0-C)

`crates/litchi-iwa/tests/exit_parity.rs` (deleted with the host) runs each
migrated operation twice on the same input bytes: once through the host
`Draft` (resolving the host's native identifiers from the selector through
host reads) and once through the leaf family. Inputs: every tracked fixture
under `test-data/iwork/**` the family applies to (native-saved Apple files
first), plus source-built packages from the leaf `Builder` (byte-identical
to the host `Builder` for the shared arguments — asserted first). For each
case it asserts:

1. **Reads.** The leaf read of both outputs and the host read of both
   outputs are equal (semantic values, order, counts), and a strict
   reopen of both succeeds through `litchi_{pages,numbers,keynote}::Package`.
2. **Bytes, tier A.** Where the host route already rewrites through the
   bounded wire layer (886 wire-patch call sites in the host), the two
   outputs are byte-identical, the whole ZIP included.
3. **Bytes, tier B.** Where the host route decodes and re-encodes a message
   with Prost (765 `encode_to_vec` sites, which drop unknown fields), the
   leaf output keeps every untouched ZIP member and every untouched object
   byte-identical to the input, and each changed object decodes (test-only
   Prost oracle) to a message equal to the host's on every known field. Any
   other difference fails the case.
4. **Identity allocation.** New object identifiers and UUIDs are equal
   (the leaf adopts the host's watermark policy), or the case records why
   not and is moved to §4.2.
5. **Refusals.** Where the leaf refuses an input the host accepted, the case
   fails unless the refusal is listed in §4.2.
6. **No-op and inverse.** An unchanged edit is an exact no-op; applying the
   inverse patch restores the input bytes.

The harness also replays the existing host test vectors (1,697 unit tests,
83 integration tests) through the leaf families where a family exists, so
the host's own assertions become leaf assertions before they are deleted.
It is kept, growing per family, until W3 deletes the host. ADR 0028 asks
for native verification only where output changes; tier A changes nothing,
so it carries the host's existing native evidence forward.

### 4.2 Planned behaviour that changes output (flagged for the owner)

None is intended, but these classes will occur and each needs D1:

| Class | Where | Expected change | Why |
|---|---|---|---|
| unknown-field preservation | every tier-B route (chart style, shape style, text-box rich text, cell rich text, builds; see the crosswalk families) | leaf output keeps fields and field order the host dropped or reordered | ADR 0028 gate 4 requires untouched bytes to be preserved |
| ZIP layout of changed members | routes where the host's `replace_archive` falls back to its logical writer (the "historical logical writer" path in `package.rs`) | compressed member bytes and local headers differ; content equal | leaf publication is exact-source reassembly |
| stricter admission | source-built or normalized legacy graphs, historical type aliases, detached models, changed Pages no-root/fallback bodies (ADR 0028: they "fail closed") | leaf refuses (`Refused`/`Limit`) where the host edited | ADR 0028 forbids post-admission legacy fallback |
| identifier allocation | duplicate/create routes if a donor allocates differently from the host | different object IDs/UUIDs, equal content | to be avoided; the harness asserts equality first |

## 5. Performance

Each WP measures, release build, pinned CPU, ABBA, on the tracked fixtures,
the host `Draft` route against its leaf family for the scenarios in the
wave table: single operation (open + edit + commit + `to_bytes`) and
1/10/100/1k-operation batches where the family batches. Bar: the leaf route
is no more than 5% (or 20 µs) slower than the host route it replaces
(performance-program regression rule), and the leaf single-op and batch
rows become the leaf baseline for the unified `File` bar of the goal. Rows
go into `tools/unified_bench_matrix.py` under new route names; the paired
bench lives in the host crate (`crates/litchi-iwa/examples/exit_bench.rs`)
and is deleted with it, its results kept as records. Allocation and heap
checks use callgrind/heaptrack (`perf` is blocked on this host). Known
risks: the donors' admission ledgers precharge work and may cost more than
the unbounded host on small edits; tier-B routes trade a Prost re-encode for
a wire rewrite, usually cheaper.

## 6. Deletion gate mapping

| ADR 0028 gate item | State today | Closed by |
|---|---|---|
| 1 no manifest depends on `litchi-iwa` | holds (§2.1) | W3 re-verifies |
| 2 every module, example, fuzz target, generated-schema/build path, fixture has a focused owner | modules mapped (§3.1); examples, tests and fixture provenance open; no host fuzz target or build script exists (`litchi-iwa-protos` owns the generated schemas) | W0-C (ledger), W2-D |
| 3 semantic parity for removed behaviour (Numbers order/orphans, Pages empty/fallback bodies, Keynote rich storage and aggregate limits) | read-side parity established by the ADR 0028 entries of 2026-08-08; the harness re-proves it on every fixture | W0-C, W1-* |
| 4 selector-first transactions preserving untouched bytes and passing native open/save/reopen | per family; native part needs D1 | W1-*, W2-* |
| 5 no root-facade re-export, feature alias or type alias | holds (`crates/litchi/Cargo.toml` has no host edge) | W3 re-verifies |
| 6 no host or host debt in the boundary policy | 11 debts open | W3 |

## 7. The unmerged Codex branches (coordinator addition)

### 7.1 What they are

`git for-each-ref refs/heads/` (excluding `uo/*`, `archive/*`, trunk and
the `uo2/*` branches of this program) lists **66** iWork branches from
2026-10-23/24. All share one merge base with trunk, **`788bc28538`**
(2026-10-13, the `feat/office-format-completeness` tip before the unified
API program), and are **4,090 commits behind** trunk. None carries a
`docs/unified-ops` record. They form one integration line and side
branches:

- **Shared base P0** (9 commits, `b257dd7ba2..c1c17b5314`): retires all 206
  host examples (−15.9k host lines), moves the native stylesheet and
  cell-value oracles to Pages, iwa-text prepared-decode admission.
- **Base P1** = `refactor/iwa-table-lifecycle-audit-20261023` (19 commits,
  tip `5818f7613f`): P0 + oracle moves (Keynote audio/media, Numbers
  comments, document readers), number-attachment lazy decode,
  `numbers-wire` bounded formula-graph remapping.
- **Integration line** = `refactor/iwa-exit-current-20261023` (54 commits,
  tip `3f5d3cb497`, the most complete branch): P1 + squashed checkpoints
  of theme codecs, Keynote chart lifecycle, native table-read oracle moves,
  physical/logical catalog preflight, fuzz-harness ownership (Numbers-wire,
  IWA-core), ZIP writer/reader admission, generated round-trip ownership,
  Keynote image lifecycle, Keynote layout reassignment, Pages body-drawable
  comments, ingress/Deflate admission, serialization reports, Numbers sheet
  duplication/deletion, text-rewrite heap admission, Pages header/footer and
  rich-text admission, identity-clone admission, **Pages and Keynote package
  builders** (`b66e1c36cb`), Numbers semantic resource admission. Its ADR
  0028 additions (806 lines, 40 sections) say each checkpoint is open:
  "Review still requires complete candidate-reopen admission…", "the native
  application gate remain[s] open before chart lifecycle ownership is
  considered" and "this does not establish new native-application open/save
  evidence". `feat/numbers-sheet-lifecycle-admission4-20261024` and
  `repair/pages-body-shape-b66-admission-20261024` point at the same commit
  `b66e1c36cb` (the second name is misleading).
- **Side branches**, forked from P1 or from points on the integration line,
  each carrying one feature not on the line (files checked against the
  `b66e1c36cb` tree).

Merging any of them as is would conflict heavily (`git merge-tree` against
trunk: 317 conflicted paths for the integration line, 111 outside
`crates/litchi-iwa/examples`; 222/16 for the conditional-highlight branch;
226/20 for the Pages media branch; 308/102 for the Keynote table branch),
and their API predates contract 0351: the sampled new families have no
`Refused`, no `Commit::into_parts`, 14–77 `*_with_limits` names per file
and `LimitExceeded` variants. They do follow selector-first shapes without
raw IDs. I did not build or test them (read-only; no record of a passing
gate exists); "review" below means what their commits and ADR text record.

### 7.2 Families, latest branch, verdict

| Family | Branches (count) | Latest / most complete | Migrates | Review state recorded | Verdict and order |
|---|---|---|---|---|---|
| shared base and integration line | `refactor/iwa-table-lifecycle-audit-20261023`, `refactor/iwa-exit-current-20261023`, `feat/numbers-sheet-lifecycle-admission4-20261024` = `repair/pages-body-shape-b66-admission-20261024` (4) | `refactor/iwa-exit-current-20261023` `3f5d3cb497` | see §7.1 | every checkpoint "open" in its own ADR text | **donor for W0-A/W0-B first** (substrate, codecs, `numbers-wire` formula graph), then the per-format checkpoints for W1-A/B/C; port, do not merge |
| Numbers sheet lifecycle | `feat/numbers-sheet-lifecycle-{focused,parity,repair,repair2,repair3,repair4,admission,admission2,admission3}-20261024` (9) | squashed into the integration line (`42989fcfb5`, later than every chain tip); `repair3`/`admission3` have 2–5 commits not on the line | `sheet_duplication.rs`: selector-first sheet duplicate/delete | repair chain, no pass recorded | **superseded** by the integration line; W1-A diffs `repair3`/`admission3` for unported fixes |
| Numbers table lifecycle and formula graph | `refactor/iwa-table-lifecycle-{fixes,native-tests}-20261023`, `refactor/formula-graph-migration-20261023`, `test/formula-owner-budget-20261023` (4) | `refactor/iwa-table-lifecycle-native-tests-20261023` `5b12fd761a` + `test/formula-owner-budget-20261023` `8743d9f106` (siblings off `accf4e036a`) | bounded table graph lifecycle (`numbers_table_lifecycle_codec`, a second `table_lifecycle.rs` that collides with trunk's record-0301 file), native round trips, owner budget test | none recorded | **donor W1-A** (merge both tips' content into trunk's existing `table_lifecycle` family, not as a second file) |
| Numbers conditional highlights | `refactor/numbers-conditional-highlight{,-review,-review-fix,-admission-review,-admission-review-r2}-*` (5) | `…-admission-review-20261024` = `…-r2` `2b0a37b4cc` | `conditional_highlight{,_codec,_native}.rs`, `edit_table_cell_conditional_highlighting` | review rounds, no pass recorded | **donor W1-A** |
| Numbers drawable comments | `feat/numbers-drawable-comments-focused{,-rereview}-20261024` (2) | `…-rereview` `4f400a6302` | `sheet_drawable_comments.rs` | "rereview" name, no pass recorded | **donor W2-A** |
| Numbers text admission | `review/numbers-text-{admission,admission-typed,shared-work}-*` (3) | WIP commits only (`10386cccb2` "WIP: preserve … pending shared work admission") | aggregate text decode admission | WIP | **superseded** by the integration line (`232c15c37f`, `e90002c3b6`) |
| Pages body shapes | `refactor/pages-body-shape-review-20261024`, `repair/pages-body-shape-a408-20261024` (2) | `repair/pages-body-shape-a408-20261024` `ece66b17ff` (one squash on the line; the review branch holds the 14-commit history) | `body_shape_lifecycle.rs`, `body_shape_properties.rs`, `shape.rs`, `shape_properties_codec` | "close … admission gaps" rounds | **donor W1-B** (shape codec to W0-B) |
| Pages media | `refactor/pages-media-lifecycle-{focused,full,admission}-20261024` (3) | `…-admission` `7df17e8661` | `body_media.rs`, `media.rs`, Numbers `media.rs`, new crate `litchi-iwa-media` | "close media operation admission gaps", no pass recorded | **donor W1-B** (Numbers part to W2-A); the new crate needs D6 |
| Pages table comments | `feat/pages-table-cell-comment-lifecycle-20261024`, `feat/pages-table-comment-{admission,ownership}-20261024`, `review/pages-table-comment-ownership-20261024` (4) | `feat/pages-table-comment-ownership-20261024` `131af0285f` (the `review/` sibling is a WIP repair of the same commit) | `body_table_comments.rs`, `body_table_comment_writer.rs` | WIP sibling, no pass recorded | **donor W1-B**; diff the `review/` sibling's `8151c775b2` (typed read failures) |
| Pages construction | `pages-construction-boundary-20261024` (1) | `ad2975ad52` (12 commits after `e90002c3b6`) | full source-free Pages builder ledger, schema-owned Buffa construction encoder | none recorded | **donor W1-B** after the integration line's `package_builder.rs` |
| number attachments | `refactor/iwa-attachment-reader-admission-20261023`, `uo2/iwa-storage-decode-admission-20261023` (2) | `refactor/iwa-attachment-reader-admission-20261023` `5fcae790d4` (13 commits not on the line; forks before P1) | Pages/Numbers/Keynote number-attachment lifecycles, `storage_text.rs` | none; `uo2/…storage-decode` is an older copy of three P0 commits | **donor W1-D** (shared) then W1-B/W1-A/W1-C families; `uo2/iwa-storage-decode-admission-20261023` **superseded** |
| Keynote chart lifecycle | `feat/keynote-chart-lifecycle-{focused,admission}`, `repair/keynote-chart-{consumer,review,clone-review,deploy}-20261024` (6) | `repair/keynote-chart-deploy-20261024` `7c8ff83a0c` (on `b66e1c36cb`; contains the consumer/review/clone-review repairs) | `slide_chart_lifecycle{,/}.rs`, `chart_semantic_census.rs` | integration-line text: native gate open | **donor W1-C**; shared chart graph clone to W1-E; the other five **superseded** |
| Keynote image lifecycle | `feat/keynote-image-lifecycle-{focused,candidate,candidate-05ce,candidate-a625,admission-final,admission-followup,detector-followup}-20261024` (7) | squashed into the integration line (`05cebb7ff6`, `1d2276361b`); `…-admission-followup` keeps 11 and `…-detector-followup` 3 commits not on the line | `slide_image_lifecycle.rs`, `slide_image_creation.rs`, `image_creation.rs` | candidate/follow-up rounds | **superseded** by the integration line; W1-C diffs the two follow-ups |
| Keynote shape and text-box lifecycle | `refactor/keynote-shape-text-{lifecycle,remediation,repair,repair-r2…r5}-*` (7) | `…-repair-20261024-r5` `01eb6ff3fe` | `slide_shape_lifecycle.rs`, `keynote_shape_creation_codec` | five repair rounds, no pass recorded | **donor W1-C** (codec to W0-B) |
| Keynote slide tables | `feat/keynote-table-lifecycle-focused-20261024` (1) | `817ae78f0d` (17 commits on `b66e1c36cb`) | `slide_table_lifecycle.rs`, `slide_table_geometry.rs`, `slide/table/geometry.rs` | "complete bounded slide-table lifecycle remediation", no pass recorded | **donor W1-C** |
| Keynote drawable comments, catalog preflight, storage shared work, fuzz ownership | `refactor/iwa-keynote-focused-drawable-comments-20261024`, `iwa-logical-catalog-preflight-20261024`, `refactor/iwa-storage-shared-work-remediation-20261024`, `refactor/{fuzz-harness,iwa-core-fuzz}-ownership-20261024` (5) | on the integration line; 1–2 commits each not on it | oracle/fuzz ownership moves, catalog preflight | none | **superseded**; W0-A/W2-D diff the residual commits |
| wire budget | `uo2/iwa-wire-budget-fix-20261024` (1) | `5753ea5058` (1 commit on `e90002c3b6`) | bounded overlay and batched rewrite work in `litchi-iwa-common::wire` | none | **donor W0-A** (small; port first) |

### 7.3 Order

1. W0-A ports `uo2/iwa-wire-budget-fix-20261024` and the integration line's
   substrate commits (`litchi-iwa-{archive,core,detect}`, `soapberry-zip`)
   onto trunk, conforming names and errors to 0351.
2. W0-B ports the codec commits (theme, shape properties, Keynote shape
   creation, number attachment, Numbers table lifecycle) and
   `numbers-wire` formula graph/owner.
3. W1 ports, per format, the integration line's checkpoints first, then the
   side-branch donors in the order of the table, each behind the parity
   harness. The P0 host-example retirement is not ported as is: W2-D retires
   each example only when its capability's leaf example exists.

No branch is a merge candidate. After W1 lands, the coordinator can move
the 66 refs to `refs/archive/*/codex-iwa-20261024` (owner's call; nothing in
this plan deletes them).

## 8. Risks and owner decisions

| # | Decision | Recommendation |
|---|---|---|
| D1 | **Native verification on Linux.** Gate 4 requires native open/save/reopen for mutation paths, and the owner rule allows output changes only with native proof. Tier-B routes (§4.1) and stricter admission (§4.2) change output. | Accept the tier-B differential (known-field equality with the host plus untouched-byte preservation) as the parity proof for routes whose host version already has native evidence, and queue a macOS run for the rest before W3. Without that acceptance, W3 waits for a macOS host. |
| D2 | **Codex branches.** 66 branches, 4,090 behind trunk, pre-0351 API, open review state. | Use them as donors per §7 (port into trunk under the current contract); do not merge; archive the refs after W1. |
| D3 | **`litchi-iwa-structured`.** No workspace consumer (record 0361a §2); holds the only cross-format aggregate snapshot and budgets. | Keep it (ADR 0028 allows it to remain until a separate, parity-proven decision); out of scope here. |
| D4 | **Neutral table vocabulary.** Pages and Keynote tables use `litchi_numbers` types through the host; a leaf→leaf edge (`litchi-pages → litchi-numbers`) is not in the canonical policy. | Move the neutral types (data formats, topology operations, `Region`, cell `Value`) to `litchi-iwa-common::table`; the three leaf crates re-export them at their current paths, so no public path changes. Glossary rows move with them. |
| D5 | **Raw physical access.** `Package::{entry_names, entry, archive, iwa_entry_names}` and `Document::{stats, message_type_summary}` have no leaf equivalent. | Move them to `litchi-iwa-archive::unstable` (explicitly unstable, never re-exported by a format facade), as ADR 0028 allows for low-level crates. |
| D6 | **New crate `litchi-iwa-media`** (Pages media donor). | Fold its content into `litchi-iwa-common::media` (values) and `litchi-iwa-archive` (asset storage) instead of adding a 65th package; otherwise record it in ADR 0024 and the boundary policy. |
| D7 | **Unified `File` wiring.** The host's capabilities are reachable today only through the unpublished host; moving them into the leaf crates does not by itself reach `litchi-ops`. | Treat `litchi-ops` dispatch arms for the new families as a follow-up program after W2 (shared dispatch files, separate hunks), so this exit does not block on it. |

Other risks: the host's stricter-admission gaps (no-root/fallback Pages
bodies, source-built packages, historical model aliases) may leave a real
capability the leaf refuses — each WP lists such cases in its record and
the owner decides per case whether a leaf route is built or the case is
accepted as a refusal; the donors' resource-admission ledgers may miss the
5% bar on small edits; Pages and Keynote tables depend on the shared table
engine in `litchi-numbers-wire`, so W1-A's engine changes must land before
W1-B/W1-C table families (sequence inside wave 1: engine commits first, or
W1-B/W1-C start with their non-table families).

## 9. Verification of this record

No Rust code changed. Besides the record and `results/0363/`, the commit
registers `results/0363` in `results/0289/bench-matrix-curated.json`
(`evidence_without_rows`) and regenerates `results/0289/bench-matrix.md`.
The crosswalk is 305,107 bytes, below the 1 MiB evidence limit; both
pre-commit size checks print nothing.

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate` | 0 (1,275 tests) |
| `python3 tools/officecli_capability_delta.py --check` | 0 |
| `python3 tools/officecli_corpus.py --check` | 0 |
| `python3 tools/unified_bench_matrix.py --check` | 0 (after the registration above) |
| `python3 tools/check_crate_boundaries.py` | 0 (lists the 11 debts) |
| `python3 tools/check_feature_matrix.py` | 1, environmental: every `cargo check --locked` leg stops with "cannot create the lock file" (`Cargo.lock` is ignored by `.gitignore` and absent from this worktree); unrelated to a docs-only change |
| `python3 tools/non_iwork_gate.py verify` | 1, pre-existing: "package 'litchi-doc' dependency 'flate2' path …/vendor/flate2 is not a workspace package"; unrelated to a docs-only change |

## 10. W3 executed (record 0429, 2026-11-05)

W3 ran as record [0429](0429-iwa-w3-delete-host.md) on branch `uo2/iwa-w3`,
brought up to trunk `4d8a3e161f`. The branch deletes `crates/litchi-iwa` and
every tool, gate and policy reference to it. Owner decision D1 asked for the
macOS queue ([`0363-mac-queue.md`](0363-mac-queue.md)) first; every row now has
Mac evidence (record [0434](0434-mac-run-3-results.md)). State of the deletion
gate (§6):

| ADR 0028 gate item | State on `uo2/iwa-w3` |
|---|---|
| 1 no manifest depends on `litchi-iwa` | holds; the crate is gone |
| 2 every module, example, fuzz target, schema/build path and fixture has an owner | exit ledger 0 host-only rows (2,117 migrated, 119 retired), checked without the crate; the protos build script no longer reads host sources; the 797 host-built macOS inputs that the Mac evidence rests on are kept outside the repository, listed in `test-data/iwork/mac-queue/host-built-inputs.tsv` |
| 3 semantic parity | every migrated route has a Tier-A/B parity case, all passing at `4d8a3e161f` before deletion; by owner decision of 2026-11-05 cases outside record 0429 §4 are dropped, not ported |
| 4 native open/save/reopen | all 767 rows have Mac evidence, 0 failed, 0 pending (0434 §2); the owner discarded the by-reference part of the two pivot rows (record 0429) |
| 5 no root-facade re-export or alias | holds |
| 6 no host or host debt in the boundary policy | holds: `migration_hosts` and `migration_debt` empty, 0 debt items |
