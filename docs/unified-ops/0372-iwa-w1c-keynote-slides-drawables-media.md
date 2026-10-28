# 0372 — W1-C: Keynote slides, drawables and media out of `litchi-iwa`

- Status: wave W1-C on branch `uo2/iwa-w1c`, fifth landing (review repair;
  implementation `db38bb09f8`; coordinator trunk audit `0944b7319f` merged
  into this work-package branch by `833be6047c`). The remaining Keynote rows
  wait on shared engines that other waves are building (§6).
- Date: 2026-10-28.
- Plan: [0363](0363-iwa-monolith-exit-plan.md) wave row W1-C (§0), API shape
  §3.2, parity §4; harness and ledger [0369](0369-iwa-exit-parity-harness-and-ledger.md);
  codecs and the D4 gap [0368](0368-iwa-w0b-codecs-formula-graph-table-vocabulary.md);
  owner decision D1 and the [macOS queue](0363-mac-queue.md).
- `performance_claim`: every measured route is 28–92% cheaper than its
  host route (callgrind, §5). The one regression found (z-order move, +9.4%)
  was fixed before landing.

## 1. What landed

| Commit | Crate | What |
|---|---|---|
| `fab10d07ab` | `litchi-keynote` | `slide::drawable::Edit::remove` culls data whose final owner was the removed drawable (`DataInfo` + `Data/` member) and releases the object watermark when the removal takes out the object holding it, as the host does (`Diagnostics::removed_data`); `Preset` gains `LeftArrow`, `RightArrow`, `DoubleArrow`, `RegularPolygon`, `Star` (host path sources; shared `PolygonSides`/`StarPoints`/`InnerRadiusRatio` re-exported from `slide::drawable`) |
| `ea3d87a07a` | `litchi-iwa` (host) | two documented host defects fixed (§3): dangling template tags after drawable removal, missing external reference after `add_slide_text_box` |
| `ebbce329e2` | `litchi-iwa` tests, ledger, queue | drawable and slide-lifecycle parity cases; harness learns the leaf's documented policies (root-preview invalidation, save-token advance, deterministic UUIDs of created objects); Keynote drawable schemas registered |
| `e2de754f3b` | `litchi-keynote` | the 0368 gap: `slide::table::number_format::{Number, DecimalPlaces, FixedDecimalPlaces, NegativeStyle, ThousandsSeparator, MAX_DECIMAL_PLACES, NumberError, At, AtError}` are now re-exports of `litchi_iwa_common::table` (`tests/neutral_table_vocabulary.rs` proves nominal identity); no workspace caller needed a change |
| `8ff06a2d5e` | `litchi-iwa` tests | paired callgrind probe `keynote_w1c_cost` |
| `3d006110be` | `litchi-iwa-archive` | `SourceCatalog::reopen` (introduced as a hidden `__from_shared_bytes_reusing`, renamed in the follow-up slice): a derived candidate reuses the base's parsed archive for byte-identical IWA members (ZIP and budgets still checked in full); parsed components are `Arc`-shared |
| `ba7155292b` | `litchi-keynote` | every structural commit reopens its candidate through the reuse path |
| `a9660091a6` | `litchi-keynote` | layout reassignment: `slide::creation::Edit::set_layout(slide, layout)` (port of host `set_slide_layout`, including the layout-media swap) |
| `4e85d30500` | `litchi-iwa` tests, ledger, queue | layout-reassignment parity on 48 layout pairs |
| `fc445c1aa3` | `litchi-keynote` | `Builder::language`, `Builder::locale` |
| `a760655e92` | `litchi-iwa` tests | builder language/locale field differential |
| `29c9cce6df`, `ecb7cfcf6c` | `litchi-keynote`, tests | slide builds: create, update, remove and read (`slide::build`), 26 constructor families |
| `e6f00e6e50`, `620e43344e`, `437c73282d`, `1acbf4c9e5`, `a7dea05f89` | `litchi-keynote`, tests | `chart::lifecycle`: kind, series direction, frame, remove, duplicate, grid replacement |
| `e17f8bf676`, `7b355eff95` | `litchi-keynote`, tests | `chart::lifecycle::Edit::add`: a source-free chart written field by field in tag order |
| `c04114289c`, `86c2da7b8f` | `litchi-keynote`, tests | `slide::image::lifecycle`: duplicate, flip, original size, properties, data replacement, image reads |
| `4a8114b408`, `caf70ad617` | `litchi-keynote`, tests | `slide::image::text`: image title and caption on the shared caption seam |
| `5943c976ed`, `33bad9f784`, `9485d716a3` | `litchi-keynote`, tests | `slide::drawable` duplicate, flip and properties for text boxes, shapes and lines; shape text set/replace/clear |
| `498940f8f3`, `e7aaa469ef` | `litchi-keynote`, tests | `slide::drawable::Edit::set_preset` |
| `4aa08c0f1a`, `95602e161b` | `litchi-keynote`, tests | line segment read and move on the W1-E line codec |
| `7795070ea7`, `1619274c84` | `litchi-keynote`, tests | `slide::drawable::Edit::add_line` |
| `2bfdf7309e`, `e3357f7baa` | `litchi-keynote`, tests | 23 whole-text character and paragraph properties of text boxes and shapes (W1-D engine): set, rewrite, reset, read |
| `b44031f527`, `6c83e9c562` | `litchi-keynote`, tests | paragraph lists (W1-D list engine) |
| `af21c5ea69`, `8eaa28bb29` | `litchi-keynote`, tests | named paragraph styles: apply, create, rename, delete, read |
| `0ea6e25f8e`, `f06ee9a689` | `litchi-keynote`, tests | hyperlinks, highlights, comments and replies (W1-D smart-field and annotation engines) |
| `9f0a2904ef`, `b205f50e4c` | `litchi-keynote`, tests | Date & Time fields and Drop Caps |
| `6060e43192`, `a59220f0bd` | `litchi-keynote`, tests | language runs |
| `ade498d1ae` | `litchi-keynote` | named-style rename helper renamed so the `fs::rename` boundary check no longer matches it |
| `bf33443e98` | `litchi-keynote`, tests | `Package::{media_assets, slide_media_assets, media_asset_data}` through a Keynote-owned `MediaAsset` handle; package lifecycle parity |
| `5004b3e019` | tests | slide text-storage and drawable-inventory parity |
| `2dbc2f59aa` | tests | the cost probe covers six more routes |
| `231c08a016`, `368ac34ae8` | `litchi-keynote`, tests | third landing: drawing and frame styles over `litchi_iwa_style::shape` (`slide::drawable::Edit::{set_shape_style, reset_shape_style, set_image_fill, add_shape_with_style, add_line_with_style}`, `Package::slide_drawable_shape_style`, `slide::drawable::shape_style`) |
| `0dac327e07`, `850aab7420` | `litchi-keynote`, tests | `Package::slide_shape_preset`; the shape list |
| `7a6128df01` | `litchi-iwa` (host) | shape title/caption registers theme edges and invalidates the slide thumbnail (the image fix of `9045c9f3f4`, applied to shapes) |
| `044ade3caa`, `e6867f9f03` | `litchi-keynote`, tests | `Package::{slide_shape_text, edit_slide_shape_text}` on the shared caption seam |
| `99cc3301cb`, `755a2598b4` | `litchi-keynote`, tests | `slide::table::cell_appearance` over `litchi_iwa_table::cell_appearance` |
| `b81571e316`, `9bd41fa35a` | tests | the cost probe covers shape fill, shape caption, text-box columns and the table cell fill |

The host keeps every route as the oracle. Host behaviour changed only to fix
the two defects of §3 and the caption seam's missing theme references
(`9045c9f3f4` for images, `7a6128df01` for shapes). No ops dispatch file changed. The crate's one new edge,
`litchi-keynote -> litchi-iwa-style`, came from trunk (`3b9b92ac03`) after
0375a moved the embedded-media feed into that crate.

## 2. Ledger

The Keynote ledger now has 331 `migrated`, 5 `retired` and 303 `host-only`
rows. Wave W1-C migrated about 298 of them: 22 in the first landing, about
223 in the second, 44 in the third (35 shape-style, column, text-layout,
preset and caption rows; 9 table cell-appearance rows), and 9 in this
continuation (table lifecycle and cell reads). W1-F and W1-G migrated some of
the other `table.*` rows (W1-G's Keynote hidden-axes facade, `4d55e0b052`,
migrated the last two). Every migrated route without host native evidence has
a `0363-mac-queue.md` row (the continuation adds the lifecycle route).

The first landing's rows:

| Host fn | Leaf destination | Parity test | Tier |
|---|---|---|---|
| `remove_slide_{text_box,image,shape}` | `slide::drawable::Edit::remove` | `keynote_slide_{text_box,image,shape}_remove_parity` | B |
| `set_slide_{text_box,image,shape}_geometry` | `slide::drawable::Edit::set_geometry` | `keynote_slide_*_geometry_parity` | B |
| `move_slide_drawable`, `set_slide_drawable_order` | `slide::drawable::Edit::move_to` | `keynote_slide_drawable_{move,order}_parity` | B |
| `slide_drawable_order`, `slide_{text_box,image,shape}_geometry` | `Package::slide_drawable_list` | `keynote_slide_drawable_read_parity` | A |
| `add_slide_text_box` | `slide::drawable::Edit::add_text_box` | `keynote_slide_text_box_add_parity` | B |
| `add_slide_shape` (all 8 presets), `add_slide_rectangle` | `slide::drawable::Edit::add_shape` | `keynote_slide_{shape,rectangle}_add_parity` | B |
| `add_slide_image` | `slide::drawable::Edit::add_image` | `keynote_slide_image_add_parity` | B |
| `duplicate_slide` | `slide::creation::Edit::duplicate_slide` | `keynote_slide_duplicate_parity` | B |
| `add_slide`, `insert_slide` | `slide::creation::Edit::add_slide` | `keynote_slide_{add,insert}_parity` | B |
| `slide_layouts`, `default_slide_layout` | `Package::{slide_layouts, default_slide_layout}` | `keynote_slide_layouts_read_parity` | A |
| `set_slide_layout` | `slide::creation::Edit::set_layout` | `keynote_slide_layout_set_parity` (48 reassignments) | B |

Rows migrated by the second landing, by family (each parity case is named
in the ledger's `parity_test` column):

| Family | Rows | Leaf destination |
|---|---:|---|
| `keynote.build` | 47 (+5 retired) | `slide::build::Edit`, `Package::slide_builds` |
| `text.character`, `text.paragraph` | 69 | `slide::drawable::Edit::{set,reset}_text_property`, `Package::slide_drawable_text_property` |
| `text.paragraph_list` | 30 | `slide::drawable::Edit` list verbs and reads |
| `drawable.shape`, `drawable.text_box_*`, `drawable.list` | 18 | `slide::drawable::Edit::{duplicate, duplicate_with_text, flip, set_properties, set_preset, set_text, set_line_segment, add_line}`, `Package::slide_drawables` |
| `drawable.image` | 16 | `slide::image::{lifecycle, text}` |
| `text.comment`, `text.highlight`, `text.hyperlink` | 16 | `slide::drawable::Edit` annotation verbs and reads |
| `text.date_time_field`, `text.drop_cap`, `text.language` | 14 | `slide::drawable::Edit` field, drop-cap and language verbs and reads |
| `chart.lifecycle` | 8 | `chart::lifecycle::Edit` |
| `text.paragraph_style` | 8 | `slide::drawable::Edit` named-style verbs and reads |
| `slide.lifecycle`, `io`, `media.asset`, `text.storage` | 13 | `Package::{from_bytes, open, save, write_to, slides, media_assets, slide_media_assets, media_asset_data, slide_drawable_list}` |

Rows migrated by the third landing:

| Family | Rows | Leaf destination |
|---|---:|---|
| `drawable.shape` (fill, stroke, shadow, effects, text layout, line endpoints, image fill, add-with-style, preset, list) | 24 | `slide::drawable::Edit::{set,reset}_shape_style`, `set_image_fill`, `add_shape_with_style`, `add_line_with_style`; `Package::{slide_drawable_shape_style, slide_shape_preset, slide_drawable_list}` |
| `drawable.shape` (title and caption) | 5 | `Package::{edit_slide_shape_text, slide_shape_text}` |
| `drawable.text_box_properties`, `text.character` (text-box columns and text layout) | 6 | `slide::drawable::Edit::{set,reset}_shape_style`, `Package::slide_drawable_shape_style` |
| `table.cell_appearance` | 9 | `slide::table::cell_appearance` |

Read cases compute the host and leaf projections inside one read and fail
on any disagreement. The harness then compares each read across the two
outputs.

## 3. Parity findings

Documented leaf policies, accepted by the harness only in their exact shape
(`Case::leaf_invalidates_root_previews`, `Case::leaf_structural_metadata`,
`archive_diff::MetadataPolicy`; the coordinator has these on trunk):

- the leaf deletes the three root previews (natively verified, ADR 0028);
- the leaf advances the save token by one and stamps touched components;
- created objects get deterministic UUIDs (the host draws random ones);
  identifiers match the host's exactly.

Per-case normalisations (in the case modules): the leaf registers a created
slide component as native Keynote does (`Slide` / `Slide-<id>` locators) and
derives its external references from the references its objects hold (the
host copies the layout's list, with two unreferenced styles and a weak
link) — the case instead checks that every cross-component reference is
registered; a fresh slide node keeps the native default for
`slide_specific_hyperlink_count`; layout reassignment with layout media
compares external references as a set because the host registers them in
`HashSet` order.

Host defects fixed (both found by the harness): removing a drawable left
`sage_tag_to_info_map`/instructional-text entries pointing at the deleted
object while its header reference was dropped; `add_slide_text_box` did not
register the text box's cross-component style reference (`add_slide_shape`
did).

Leaf changes made to reach parity: data culling and watermark release on
removal (as the host), the five missing shape presets, `natural_size` from
the image's pixels (the host case passes the same value).

Second-landing findings:

- Image caption seam: the host left the caption's theme dependency
  unregistered. The defect was fixed in the host (`9045c9f3f4`), and the leaf
  registers it the same way.
- Image captions: the core requires new references to form a suffix of the
  header list, so the leaf adds the new graph root and then moves it into
  the old edge's slot, which is where the host writes it.
- Byte parity of authored payloads (charts, colors, series styles) needs
  fields in tag-number order, which is how prost encodes them.
- Title and body placeholders are written through `Package::edit_slide_title`
  and `edit_slide_body`, which also invalidate the slide node's rendered
  preview. The host's `set_slide_text_storage` leaves that preview stale, so
  the storage cases run on text boxes. That invalidation is the slide-text
  family's documented policy.
- `archive_diff` now also applies the metadata policy to a registry the
  host left untouched (`d57a6bfb26`).

Third-landing findings:

- Keynote theme shadows: on most decks (`chart-data-native.key`,
  `basic.key` and others) the theme's disabled shadow still carries
  enabled-shadow fields. The host's shadow routes refuse it ("disabled iWork
  shadow unexpectedly retains enabled-shadow properties"), and the leaf
  refuses it the same way through the shared decoder
  (`Unsupported::ShapeStyle`), for reads and writes. The behaviours match,
  so no mac-queue row is needed. Shadow parity runs on
  `chart-grid-authored-native.key`, whose theme shadow both sides read. Shapes
  whose style inherits such a shadow stay a known gap (§6).
- Shape and line captions reuse the image and movie caption seam. A
  ShapeInfo's drawable is the super of its ShapeArchive, so the movie edge
  codec reads the nested ShapeArchive and needs no new codec. The new graph
  root keeps the old edge's header slot, as for images.
- Cell-appearance edits with several operations publish one candidate per
  operation, as the host applies them. Each candidate advances the source's
  save token (`Registry::advance_from`), so the transaction advances it once.

Divergence found and not migrated: `slide::drawable::Edit::set_table_cell`
against the corrected host cell writer (W1-B, on trunk) on
`table-discovery.key`: the leaf leaves the column-header bucket's
`number_of_cells` stale on a clear and writes different tile offset tables
than the host on a set. The table-cell rows stay `host-only`; the fix
belongs to the shared table engine (W1-A, `litchi-numbers-wire`). The cases
are kept out of the tree until then.

## 4. API changes

- `slide::drawable::Preset::{LeftArrow, RightArrow, DoubleArrow,
  RegularPolygon { sides }, Star { points, inner_radius }}`; re-exports
  `slide::drawable::{PolygonSides, StarPoints, InnerRadiusRatio}` (shared
  `litchi_iwa_common::shape::path` values).
- `slide::drawable::Diagnostics::removed_data`.
- `slide::creation::Edit::set_layout`.
- `Builder::{language, locale}`; `builder::Error::{InvalidLanguage,
  InvalidLocale, Locale}`.
- `slide::table::number_format` and the crate root: the Number format and
  coordinate names are the shared types (breaking only for code matching
  `NumberError`/`AtError` variants exhaustively; none in the workspace).
- `litchi_iwa_archive::SourceCatalog::reopen(&self, candidate)`
  (public, documented; physical-layer reopen of a derived candidate).

- Second landing: `slide::build`, `chart::lifecycle` (`add`, `remove`,
  `duplicate`, `set_data`, kind, direction, frame), `slide::image::{lifecycle,
  text}`, the `slide::drawable::Edit` verbs listed in §2 with their reads,
  and `slide::media::assets::{MediaAsset, Kind, Error}`. A media asset is an
  opaque handle. Native data identifiers and `litchi_iwa_common` types stay
  out of the public API (`check_crate_boundaries`).

- Third landing: `slide::drawable::shape_style::{Property, PropertyKind,
  ImageFill}` with the shared drawing vocabulary re-exported;
  `slide::drawable::Edit::{set_shape_style, reset_shape_style, set_image_fill,
  add_shape_with_style, add_line_with_style}`; `Package::{
  slide_drawable_shape_style, slide_shape_preset, slide_shape_text,
  edit_slide_shape_text}`; `slide::table::cell_appearance` (`Edit`, `Patch`,
  `Commit`, `Appearance`, `Package::{slide_table_cell_appearance,
  edit_slide_table_cell_appearance, apply_slide_table_cell_appearance}`);
  `slide::drawable::Unsupported::ShapeStyle`. An image fill reads back as an
  `ImageFill` (layout and colors), never as a native data identifier.

- Fifth landing: the shared `litchi-iwa-table::structure::TableLeaf` reopen
  seam returns `FaultResult<Self>`, so candidate reopen and exact-patch replay
  preserve typed payload/selection/package limits and allocation failures.
  The Numbers and Keynote adapters retain their package-specific resource
  mapping at that seam. Keynote rejects duplicate requests for tables with
  existing merges as a typed `UnsupportedDependency` before cloning; a
  successful unreadable merged-table candidate is therefore not admitted.
  Removal parity pins the expected orphan object IDs and member names per
  fixture, verifies the component/member locator closure, known fields,
  references and untouched bytes, and includes a negative control for an
  unrelated survivor. The default archive comparison remains strict.

The first landing's glossary profiles passed unchanged: no new banned names
and no new family. `set_layout` joined the registered `slide::creation`
family. Each second-landing family is registered in `tools/api_glossary.json`.

## 5. Performance

Callgrind, release with debug info, CPU 3 pinned, `taskset -c 3`, ABBA
(host, leaf, leaf, host) in one run; inclusive Ir of `w1c_host_route` and
`w1c_leaf_route` (two calls each: open + one operation + commit +
serialize); selectors resolved outside the measured functions.
`crates/litchi-iwa/tests/parity/keynote_w1c_cost.rs`, command in its header.

| Route (fixture) | Host Ir | Leaf Ir | Δ |
|---|---:|---:|---:|
| text-box remove (`basic.key`) | 248,984,824 | 71,864,293 | −71.1% |
| image remove (`image-adjustments-native.key`) | 896,463,307 | 77,993,705 | −91.3% |
| text-box geometry | 245,212,795 | 71,305,202 | −70.9% |
| image geometry | 207,141,108 | 79,763,614 | −61.5% |
| z-order move | 110,214,034 | 79,761,714 | −27.6% |
| text-box add | 350,816,338 | 75,095,247 | −78.6% |
| shape add | 316,157,625 | 75,258,931 | −76.2% |
| image add | 498,526,155 | 76,133,732 | −84.7% |
| slide duplicate | 268,077,876 | 85,537,427 | −68.1% |
| slide add | 647,829,045 | 224,380,200 | −65.3% |
| layout reassignment (to "Title & Photo") | 1,115,466,103 | 83,864,581 | −92.5% |

Second landing (`2dbc2f59aa`, same method, `text_box_remove` re-measured
as the control: 248,624,844 / 71,874,210, −71.1%):

| Route (fixture) | Host Ir | Leaf Ir | Δ |
|---|---:|---:|---:|
| text set (`basic.key` text box) | 108,993,937 | 71,575,548 | −34.3% |
| text-box duplicate | 349,229,239 | 74,133,313 | −78.8% |
| line add | 330,487,679 | 75,139,401 | −77.3% |
| hyperlink add | 249,579,847 | 71,879,996 | −71.2% |
| text style (text property) | 467,498,853 | 123,319,648 | −73.6% |
| image duplicate (`image-adjustments-native.key`) | 452,448,983 | 81,959,059 | −81.9% |

Third landing (`b81571e316`, `9bd41fa35a`, same method; the shape routes run
on a host-added rectangle on `basic.key`):

| Route (fixture) | Host Ir | Leaf Ir | Δ |
|---|---:|---:|---:|
| shape fill set | 367,757,259 | 122,341,999 | −66.7% |
| shape caption create | 605,529,588 | 190,552,808 | −68.5% |
| text-box columns set | 398,367,503 | 122,301,083 | −69.3% |
| table cell fill set (`slide-table-read-native.key`) | 1,044,210,750 | 281,599,637 | −73.0% |

Fourth continuation (branch `uo2/iwa-w1c`, same method; two calls per route in
the ABBA probe):

| Route (fixture) | Host Ir | Leaf Ir | Δ |
|---|---:|---:|---:|
| table duplicate (`table-discovery.key`) | 2,545,874,716 | 320,310,251 | −87.4% |
| table remove (`table-discovery.key`) | 2,077,914,520 | 271,030,659 | −87.0% |

Fifth landing review repair (`db38bb09f855a6fd26b17e89e6a0eb9474a5316f`,
same release/debug-info, pinned-CPU and ABBA method):

| Route (fixture) | Host Ir | Leaf Ir | Δ |
|---|---:|---:|---:|
| table duplicate (`table-discovery.key`) | 2,546,163,008 | 391,866,125 | −84.6% |
| table remove (`table-discovery.key`) | 2,077,820,837 | 271,058,460 | −87.0% |

The release build was produced by
`CARGO_BUILD_JOBS=4 cargo test -p litchi-iwa --test parity --all-features --release --no-run`
(exit 0). The measured binary was
`target/release/deps/parity-25144fe4e83e134e`, 38,223,240 bytes, SHA-256
`c48a51619a304271f0bd9072410acf8e32e348dd3d4af948617824cfa08b93b2`.
The duplicate route command was
`W1C_ROUTE=table_duplicate taskset -c 3 valgrind --tool=callgrind --callgrind-out-file=/tmp/iwa-w1c-table-duplicate-20261028-v2.callgrind target/release/deps/parity-25144fe4e83e134e --ignored --exact keynote_w1c_cost::keynote_w1c_route_cost_abba --test-threads=1`
(exit 0); its outside-repository artifact is 493,342 bytes with SHA-256
`6c96c1ea5b4016ff5c8a3e2d134da0b55a8944f2513f6ca353c29f2584838de0`.
The remove command was the same with `W1C_ROUTE=table_remove` and
`/tmp/iwa-w1c-table-remove-20261028-v2.callgrind` (exit 0); its artifact is
443,491 bytes with SHA-256
`7e7fe92251fff900f245a558aa69236872478333815d22d90a9e4b7da1411e17`.
Both results were extracted with
`callgrind_annotate --inclusive=yes /tmp/<route>.callgrind | rg 'w1c_(host|leaf)_route|PROGRAM TOTAL'`.
The artifacts are intentionally not tracked. These branch-local numbers must
be repeated on the final integrated trunk candidate.

These figures are branch-local evidence; the coordinator must repeat both
routes on the final integrated trunk candidate before treating them as the
qualification result.

Before `3d006110be`/`ba7155292b` the z-order move was **+9.4%** (leaf
120,550,699 vs host 110,203,188): the structural reopen parsed every
component again (22.5M Ir of the 60M leaf route). Reusing the base's parsed
archives for unchanged members fixed it and cut every structural route by a
further ~40M Ir.

## 6. Gaps and follow-ups

The remaining 312 Keynote `host-only` rows all wait on shared engines or
harness work that other waves own. The coordinator ruled that W1-C must not
build its own copy of any of them.

| Gap | Rows | Owner and next step |
|---|---:|---|
| Chart style families | 160 | W1-E/W2 `litchi-iwa-style`; Keynote facades after it |
| Table rich text, formats, comments, conditional highlights, topology | 127 | W1-G `litchi-iwa-table`; thin Keynote facades after each group |
| Remaining table lifecycle inventory (`slide_tables`, `slide_table`, `add_slide_table`, `set_slide_table_geometry`, `resize_slide_table`) | 5 | W1-C follow-up; the continuation's cell projections and duplicate/remove route are landed, while add/geometry/resize still need a complete shared graph policy |
| Builder rows (including `language`/`locale`, implemented in `fc445c1aa3`) | 10 | the host builder draws random UUIDs; the rows wait for the harness to normalise builder identities |
| `replace_media` | 1 | the media-replacement transaction (`slide::media::data`) covers movies; image and audio routes join after W1-E's media engine |
| Shapes whose style inherits a disabled theme shadow that keeps enabled-shadow fields (most decks) | — (known gap, inside migrated rows) | both the host and the leaf refuse their shadow reads and writes (§3); supporting them needs a native-evidence decision on how such a shadow reads |

## 7. Verification

Final run of the third landing. The cargo gates ran at `9bd41fa35a` (trunk
`d6b7f850cf` merged). After trunk `dc9902fe5a` was merged, clippy, the
Keynote parity cases (238), the `litchi-keynote` tests and every Python gate
were re-run. Build environment from the coordinator rules
(`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`).

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-keynote -p litchi-iwa-archive -p litchi-iwa -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings`, and with `--features keynote` | 0, 0 |
| `cargo test -p litchi-keynote --all-features` (1,194) / `-p litchi-iwa-archive` (215) / `-p litchi-iwa` (2,402; `parity` target 609 passed, 58 ignored probes) / `-p litchi` (209) | 0 / 0 / 0 / 0 |
| `cargo test -p litchi-ops --all-features`: `--lib`, `--doc`, nine groups of ≤20 `--test` targets | 0 each |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-keynote -p litchi-iwa-archive` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (73 profiles) | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0; 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_iwa_exit_ledger` | 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |
| `python3 tools/officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `unified_bench_matrix.py --check` | 0, 0, 0 |
| `python3 tools/check_feature_matrix.py` | 0 |

`litchi-ops` capability and corpus behaviour did not change: no ops
dispatch arm was added (D7), so the goldens and the OfficeCLI corpus were not
re-blessed. Both pre-commit size checks printed nothing for every commit.

Continuation scoped verification (source `db38bb09f855a6fd26b17e89e6a0eb9474a5316f`;
all commands below returned exit 0; test counts are included where applicable):

| Check | Result |
|---|---|
| `cargo fmt --all -- --check` | 0 |
| `CARGO_BUILD_JOBS=4 cargo test -p litchi-keynote --test slide_table_lifecycle --all-features -- --nocapture` | 4 passed |
| `CARGO_BUILD_JOBS=4 cargo test -p litchi-iwa --test parity --all-features keynote_slide_table_remove_divergence_is_reported -- --nocapture` | 1 passed |
| `CARGO_BUILD_JOBS=4 cargo test -p litchi-iwa --test parity --all-features keynote_slide_table -- --nocapture` | 38 passed, 0 failed, 739 filtered |
| `CARGO_BUILD_JOBS=4 cargo test -p litchi-numbers --test table_relocation --all-features -- --nocapture` | 9 passed |
| `CARGO_BUILD_JOBS=4 cargo test -p litchi-numbers --test table_cell_number_format --all-features -- --nocapture` | 14 passed |
| `CARGO_BUILD_JOBS=4 cargo clippy -p litchi-keynote -p litchi-numbers -p litchi-iwa-table --all-targets --all-features -- -D warnings` | 0 (vendor warnings only) |
| `RUSTDOCFLAGS=-Dwarnings CARGO_BUILD_JOBS=4 cargo doc --no-deps --all-features -p litchi-keynote -p litchi-numbers -p litchi-iwa-table` | 0 (vendor warnings only) |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0; 73 profiles, 3,957 banned names, 704 families, 125 immutable owners |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0; 70 packages, 313 declarations, 12 explicit debt items |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `CARGO_BUILD_JOBS=4 cargo check -q --locked --manifest-path crates/litchi-keynote/fuzz/Cargo.toml --bins` | 0 |
| `CARGO_BUILD_JOBS=4 cargo check -q --locked --manifest-path crates/litchi-numbers/fuzz/Cargo.toml --bins` | 0 |

The fuzz-lock commands were run independently and their cargo statuses were
captured directly; no grep-filtered pipeline was used as a success signal.

## 8. Fourth landing — slide-table lifecycle and cell reads (continuation)

The fourth landing ports the bounded Keynote table lifecycle operations that
can be expressed through the shared `litchi-iwa-table::structure` engine.
`Package::edit_slide_table_lifecycle` resolves a `SlideSelector` and
`TableSelector`, stages exactly one `duplicate` or `remove`, publishes an
exact-source candidate, strictly reopens it, and verifies the selected table
graph before returning `Commit`. `Patch::inverse` restores the retained source
bytes exactly; replay on another source is a typed `PatchConflict`. A semantic
or directory-backed source is refused as `UnsupportedSource`, and malformed or
ambiguous rooted graphs never fall back to the host.

The duplicate facade uses the shared graph clone and formula-owner remapping,
then attaches the new `TableInfo` to the slide's owned-drawable and z-order
lists. It keeps the host's copy-name and ten-point offset rules. Removal uses
the shared private-closure and formula-owner planner, detaches template-tag
and instructional-text references, releases a removed identifier suffix, and
proves that no removed object remains reachable. The host's known orphan-island
behaviour after `remove_slide_table` remains an explicit Tier-B divergence;
the leaf's complete private-closure removal is intentional and documented in
`keynote_slide_table_remove_divergence_is_reported`.

The archive-free `Package::slide_table_cells` and
`Package::slide_table_merges` reads now have differential coverage over the
host cell count, comment count, cell/comment accessors, iterators and merge
projection. The native-authored read receipt proves the read route only; the
two mutation routes remain in the macOS queue pending native open/save/close/
reopen evidence. `add_slide_table`, geometry and resize stay host-only because
their current host scaffolds and shared engines do not yet admit a complete
selector-first facade.

The parity cases cover `table-discovery.key`, `slide-table-read-native.key` and
`slide-table-merges-native.key`: duplicate is Tier B with documented preview,
save-token and fresh-UUID policies; reads are Tier A against the native read
receipt; removal is Tier B with the intentional orphan-island divergence.
The release ABBA cost probe adds `table_duplicate` and `table_remove` routes to
`keynote_w1c_cost.rs`; final instruction-count values are recorded after the
coordinator's integrated-tree qualification because later merges invalidate
earlier measurements.

`slide_table_lifecycle::low_selection_budget_is_typed_and_atomic` opens the
same fixture with `max_references(1)`, observes the selector's typed
`LimitKind::Selection` refusal, and confirms that the exact source bytes are
unchanged. The selector verification and exact-patch reopen paths preserve
selection, payload and package limits plus allocation failures rather than
collapsing them into `Verification`; removal-side vectors reserve through the
shared table admission helper before any growth.

The ledger moves `duplicate_slide_table`, `remove_slide_table`, the seven cell
read accessors and `merges` to their leaf destinations. The lifecycle mutation
route is queued in [0363-mac-queue.md](0363-mac-queue.md); read evidence cites
`test-data/iwork/keynote/slide-table-read-native-receipt.json`. No `litchi-ops`
dispatch or capability schema changed; that remains D7.

### Review repair in the fifth landing

The shared table publisher now exposes `TableLeaf::reopen_candidate` as a
`FaultResult<Self>`. `Working::publish` and the exact-patch replay path pass
that result through without converting resource failures to `Verification`.
Numbers maps archive/common/package limits to its established package
resource kinds (including `Allocation`), and Keynote applies the equivalent
mapping for candidate ingress. This is the seam that the Pages adapter must
adopt when its work-package branch is integrated; it introduces no alias or
duplicate public API.

Keynote table duplication now inspects the selected table's merge projection
before cloning. Any existing merge is refused with typed
`UnsupportedDependency`, because the former successful candidate could not be
read by either table reader. The merged-table parity case asserts this refusal
and leaves the source bytes unchanged.

Removal parity no longer derives an allowance from the observed host-minus-
leaf set. Each tracked fixture has a pinned orphan object-ID set and pinned
host member list; an allowlisted member is accepted only when its locator maps
to the corresponding component, and the component/member closure is removed
from the normalization comparison. All surviving known fields, references,
object identities and untouched members remain strict. A negative control
removes an unrelated surviving drawable and asserts that the same fixed
allowance still reports a violation. The no-policy/default archive helper
remains strict. Shared admission reserves all removal-side vectors before
growth, and the low-budget lifecycle case records a typed refusal with exact
source preservation.

There is still no native mutation open/save/close/reopen receipt for these
routes, so the queue entry remains pending and the evidence is Tier B only.
