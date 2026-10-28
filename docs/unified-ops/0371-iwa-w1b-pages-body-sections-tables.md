# 0371 — W1-B: Pages body media, shapes, text fields and table cells out of the host

- Status: resumed merge-ready slice on branch `uo2/iwa-w1b` (the branch was
  merged with coordinator trunk `7973aef0e6` before this handoff). The wave
  continues (§6).
- Date: 2026-10-28 (record opened 2026-10-26).
- Plan: [0363](0363-iwa-monolith-exit-plan.md) wave W1-B (§0, §3.2, §4, §7);
  harness and ledger [0369](0369-iwa-exit-parity-harness-and-ledger.md);
  shared engines from W0-B [0368](0368-iwa-w0b-codecs-formula-graph-table-vocabulary.md),
  W1-A [0370](0370-iwa-w1a-numbers-tables-cells.md), W1-C
  (`SourceCatalog::reopen`) and W1-D (record 0374: `litchi_iwa_text_wire::{ranged,
  smart_field}`); owner decision D1 and the [macOS queue](0363-mac-queue.md).
- `performance_claim`: every migrated route is at or below its host route
  under callgrind ABBA (§4); the one miss found (whole-body set, +9.1%) was
  fixed in the leaf, then all routes gained from the shared candidate reopen.

## 1. What landed

| Commit | Family / change | Host rows |
|---|---|---|
| `95efcc2b40`, `2a44d21df3` | `litchi_pages::drawable::properties` — `Package::{body_drawable_properties, edit_body_drawable_properties, apply_body_drawable_properties}` over the new `DrawableSelector::{Image, Movie, Audio, Shape}` (kind + text-flow position, the 0363 §3.2 Pages drawable selector); `Edit::{set_geometry, set_position, flip, restore_original_size, set_properties, set_playback}`; values `drawable::{Geometry, GeometryFlags, Properties}`. The shared private proof `package::body_drawable_graph` resolves the rooted body, its attachment table and the ownership of each drawable, and publishes in-place message rewrites | 25 (image/movie/audio/shape geometry, flip, Original Size, drawable properties, audio position, movie/audio playback) |
| `06c52406c2` | host `set_body_text`/`clear_body` → `Package::edit_body_text`; `replace_body_text` (body-absolute range, never across a section break) → a section-relative `Package::edit_section_text` replace | 3 |
| `db808798af` | `litchi_pages::drawable::text` — `Package::{body_drawable_text, edit_body_drawable_text, apply_body_drawable_text}`: the plain text an ordinary body shape owns, one TSWP splice per edit; a splice that would drop a referenced object is refused | 3 |
| `907f3afcf2` (and `e805eca0f5` on `uo2/iwa-w1b-cellwriter`, merged to trunk) | **host defect fixes** in the shared TST cell writer (`crates/litchi-iwa/src/numbers/editor/{storage,model}.rs`): tile `numrows` keeps the native row-count convention on cell writes (record 0370's helper), and a cell write adjusts the column-header cell count as native saves do (`body-table-cells-native.pages` counts both axes); macOS queue rows for the Numbers, Keynote and Pages host set-cell routes | — |
| `b8599dfd3b`, `96bdc43cd2`, `1d0ddd2e3e` | host `set_table_cell`/`set_table_cells`/`clear_table_cell` → `Package::edit_body_table_values` (the clear after W1-A's writer and host fix); the cell, table-list, merge and comment reads (`tables`, `table`, `get_cell`, `iter_cells`, `cell_count`, `get_comment`, `iter_comments`, `comment_count`, `merges`, `table_cell_comment[_replies]`) → `body_tables`, `body_table_cells`, `body_table_merges` | 14 |
| `a408e00fbd`, `9bac132398` | `litchi_pages::bookmark` — `Package::{body_bookmarks, edit_body_bookmarks, apply_body_bookmarks}`; `Edit::{add, update, remove}` by position or name; on W1-D's `ranged` and `smart_field` engines | 4 |
| `c0ec0bac27`, `9bac132398` | `litchi_pages::date_time` — `Package::{body_date_time_fields, edit_body_date_time_fields, apply_body_date_time_fields}`; `Edit::{add, insert, update, remove}` (insert adds the display text) | 5 |
| `a9d2af35ca`, `de1dbad71e` | `litchi_pages::number_attachment` — `Package::{number_attachments, edit_number_attachments, apply_number_attachments}` over `Target::{Body, HeaderFooter(HeaderFooterSelector)}`; `Edit::{insert, update, remove}`; on W1-D's `number_attachment` engine | 6 |
| `658c8bb696` | `litchi_pages::drawable::caption` — `Package::{body_drawable_caption, edit_body_drawable_caption, apply_body_drawable_caption}` over `DrawableSelector` + `Kind::{Title, Caption}`; `Edit::{set, clear}`. A present label's storage text is rewritten; an absent label gets the native inline graph from the shared `keynote_chart_caption_graph_codec` (style, CaptionInfo, storage, placement; the payloads the host writes); clearing points the drawable at a fresh stand-in and keeps the old graph. New objects get deterministic UUIDs in the document component and raise the watermark (`body_storage_objects::Metadata::registered`) | 15 |
| `b8116873a5` | **host defect fix**: Pages number-attachment removal deletes the removed object (§3, item 3); body and header/footer removal → `Package::edit_number_attachments` `Edit::remove` | 2 |
| `26ce96d0f7` | parity for existing leaf families: host `append_section`, `remove_section`, `sections` → `Package::{edit_section_lifecycle, sections}` (tier B); host `remove_table` (already delegating) → `Package::edit_body_table_deletion` (tier A) | 4 |
| `a71731f9f2` | `litchi_pages::drawable::lifecycle` — `Package::{edit_body_media, apply_body_media}`; `Edit::remove(DrawableSelector)` for images, movies and audio: deletes the U+FFFC and attachment entry, the drawable's z-order entry, the drawable/label/attachment graph, its data-reference owners, UUIDs and watermark suffix, and every data asset nothing references any more (`Data/` member and `DataInfo`); a commented drawable is refused. **Host defect fix**: media removal prunes the deleted attachment from the body storage header (§3, item 4) | 3 |
| `24f47b31ea` | `drawable::lifecycle` `Edit::duplicate(DrawableSelector, TextPosition)` for images, movies and audio: the host's graph clone (fresh identifiers in graph order, reference remaps by message type, CaptionInfo through the shared codec), the 12 pt (audio 30 pt) frame and attachment offset, the new U+FFFC and attachment entry, the z-order append, UUIDs and shared data-reference owners | 3 |
| `80cba9916f` | `drawable::lifecycle` `Edit::{replace_data, replace_poster, replace_asset}` (the `Data/` member and its `DataInfo` digest and materialized length, identifier kept) and the reads `Package::{media_assets, section_media_assets, extract_media}` on W1-E's embedded-media inventory, with the host's reachability from the document root or a section | 8 |
| `6c58f19db1` | `Package::body_media` → `drawable::lifecycle::Media` (selector, anchor, data/thumbnail/poster assets, original and natural sizes; frame, properties and playback through `body_drawable_properties`) replaces the host's `body_{images,movies,audio}` | 3 |
| `d67dd4a62a` | `drawable::lifecycle` `Edit::{add_image, add_movie, add_audio}`: the host's typed graph (drawable with the theme's first image or movie style preset, two caption stand-ins, attachment), the drawable z-order created when the document has none, new `Data/` members and `DataInfo` records (`stem-identifier.extension`), UUIDs, data-reference owners, the stylesheet external reference and the watermark | 3 |
| `84edbd367d` | `section::lifecycle` `Edit::insert(source, TextPosition, name)`: a section break anywhere in the body, the new section copying the source section's masters (the host's `insert_section`) | 1 |
| `2365445db4`, `f44c4ea905` | `Package::body_storage` → `body::Storage` (the shared rich-text `litchi_iwa_text::storage::Storage`, re-exported as `litchi_pages::body::{Storage, Run}`): the rooted body story with its complete native text and runs, replacing the host reads `body_storage` and `body_text` (`Storage::text`). `f44c4ea905` moved the read out of the section-text owner, whose boundary forbids storage types and the flat `body_text` name | 2 |
| `b5af407b34` | `table::sort::execution` — `Package::{edit_body_table_sort_execution, apply_body_table_sort_execution}`; `Edit::{apply, apply_to_rows}`: the host's scalar executor at wire level (stable permutation by the shared planner; tile row records, row header buckets, row UID map and stroke layers move; user-hidden rows restored to their positions in the active hidden-state extent). The payload rewrites take only payloads and a permutation, for the shared table crate | 2 |
| `8da58dfbfe` | `table::size` — `Package::{edit_body_table_size, apply_body_table_size}`; `Edit::set(rows, columns)`: the host's resize (tiles trimmed, header buckets kept inside the extent and rewritten with explicit cell counts, UID-map identities appended or dropped, stroke-sidecar and model counts); after W1-A's host fix (record 0370e) also the table's own formula-owner table and body extents (`3d59e7cbc8`) | 1 |
| `01ebb98c65` | `table::creation` (record 0329) at host parity, on W1-A's aligned `litchi_numbers_wire::table_create` (record 0370f): the host's identifier order (drawable, model, owned storage in source order, formula owner, attachment), UID maps seeded with the new model's identifier, sorted model references (the host's `table_model_references`), the drawable declaring its parent and model, the template drawable's frame kept, the formula owner's UUID registered, registrations in identifier order. Host `add_table` → `Package::edit_body_table_creation` (tier B) | 1 |
| `45734adda6`, `b24c0c46c2` | `table::cell_appearance` — `Package::{body_table_cell_appearance, edit_body_table_cell_appearance, apply_body_table_cell_appearance}`; `Edit::{set_fill, reset_fill, set_layout, reset_layout}`; `Appearance { fill, layout }`. The format-neutral engine is the new `litchi_iwa_table::cell_appearance` module (the host's copy-on-write cell-style graph: private variations in the stylesheet, the style list's key allocation and refcounts, in-place rewrites of a cell's own variation, removal of a variation that becomes empty); the Pages owner applies its `Plan` and the registry changes. Image fills are refused (data-reference bookkeeping not ported) | 6 |
| `613d5e28c9`, `287c252b4a` | cell borders: `litchi_iwa_table::cell_appearance::{borders, plan_border}` (the host's stroke layers: the highest-order run covering the cell on each side; a set restrokes the cell's own one-cell run in place, appends a run to the side's layer at the cell's index or creates that layer, with the next order); `Edit::{set_border, clear_border}` and `Appearance::borders` replace `{set,clear}_table_cell_border` and `table_cell_borders` | 3 |
| `689b1cd6d6` | `litchi_pages::builder::Builder::{language, locale}` (the host's fields: the document language and locales, the settings' language and creation/last locales, every storage language table; `PackageError::{InvalidLanguage, InvalidLocale}`); `Identity` gains its version UUID (`from_parts`, `version_uuid`), written to `versionUUID`, the properties' `revision` and the package metadata revision, as the host does — before, every leaf-built document carried the template's version UUID. Identity, revision, language, locale and the body text are one publication (the body-text transaction only validates) | 5 (`new`, `body_text`, `language`, `locale`, `build`) |
| `923ee3148b` | `drawable::style` — `Package::{body_drawable_style, edit_body_drawable_style, apply_body_drawable_style}`; `Edit::{set, reset}` over `Property::{Fill, Stroke, Shadow, Effects, TextLayout}`: the shared `litchi_iwa_style::shape` engine (W1-E, record 0375a) on a workspace over this package's parsed components, one publication per edit | 15 |
| `f9ebce8575`, `27619e3f3c` | shared `litchi_iwa_style::line_end` codec and `drawable::style` line-endpoint facade — `PropertyKind::LineEndpoints`, `Endpoints` and `Endpoint`; `Edit::{set, reset}` and `Package::body_drawable_style` replace host `set_body_line_endpoints`, `reset_body_line_endpoints` and `body_line_endpoints`, with fresh variation UUIDs on allocation and exact in-place rewrites/resets | 3 |
| `7a1b5c4054` plus resumed bounded/refusal hardening | `table::creation::Edit::duplicate(BodyTableSelector, Position)` over `litchi_iwa_table::structure::duplicate`: copies a populated body table's cells, styles, formula lists and owner family, allocates a new drawable/model/storage graph, names it with the next free `"<name> copy"`, anchors its attachment in the body story, and publishes the registry references; typed selector, dependency and position refusals remain terminal. A source with non-empty merged-cell geometry is preflighted through the bounded merge reader and terminally refused as `UnsupportedDependency` until the shared clone can rebind its inline merge owner. | 1 |
| `acf5fd71ad` | crate-boundary gate: the bookmark, date-time and attachment re-export modules move to their own files; the movie-playback owner keeps its feature-gated declaration with a `#[path]` twin for builds without the host seam; the retired-host-route audit skips the parity harness (whose cases call host routes by design) | — |
| `234929c79b`, `28fb68c468` | performance (§4): the section-text preflight sums retained header lengths instead of re-encoding the archive and clones the parsed component instead of reparsing it; every Pages commit-time candidate reopen calls `SourceCatalog::reopen` | — |

A private engine `package::body_storage_objects` holds what creating and
deleting a storage-owned object needs (the host's rules): the identifier one
above every physical object and metadata registration, the metadata
watermark raise, external-reference and UUID-registration removal and the
watermark-suffix release, and the TSWP text splice.

In all, the first slice flipped 60 Pages ledger rows to `migrated` (63
with W0-C's merge rows); with the later slices the ledger lists 114
migrated and 553 host-only Pages rows. Every new public type follows ADR 0034; the families are
registered in the `pages` glossary profile (`drawable::properties`,
`drawable::text`, `bookmark`, `date_time`, `number_attachment`) and listed
in `docs/API_GLOSSARY.md` under record 0371. No crate edge, manifest or
litchi-ops dispatch changed (D7).

The host routes stay the oracle (they are not rewired to delegate): each
host route keeps its `Draft` mutation model, and the leaf output is already
byte-identical for the tier-A routes.

## 2. Parity

Cases live in `crates/litchi-iwa/tests/parity/pages_{drawable,text,table}_cases.rs`.

| Route family | Cases | Tier | Native evidence |
|---|---|---|---|
| drawable properties (image ×4, movie ×5, audio ×3, shape ×3) | 15 + one read sweep over every media and shape | A | audio playback: `pages/audio-playback-native-resaved.pages`; others queued |
| whole-body text | 3 | A | queued |
| shape text | 3 | A | queued |
| table values and reads | 4 + one read sweep over four native fixtures | A | reads: `body-table-cells-native-receipt.json`; writes queued |
| bookmarks | add (B), update (A), remove (A) | A/B | queued |
| Date & Time fields | add (B), insert (B), update (A), remove (A) | A/B | queued |
| number attachments | body insert/update/remove, first-header insert/update/remove | A | queued |
| titles and captions (image, movie, shape × title, caption) | set: create on an unlabelled drawable, then replace on a host-labelled one; remove on a host-labelled one | B (fresh UUIDs; the replace runs are byte-identical) | queued |
| section append/insert/remove, sections read | 3 | B (UUIDs, save token) | queued |
| body-table remove | 1 | A | queued |
| table sort execution (keys; borders; hidden axes; selected rows) | 2 | A | queued |
| table resize (grow, shrink) | 1 | A | queued |
| cell fill and layout set (new variation; two cells in one edit) | 3 | B (the variation's fresh registry UUID) | queued |
| cell fill and layout reset, in-place rewrite, variation removal | 4 | A | queued |
| cell fill, layout and border reads (every cell of four native fixtures and a host-styled one) | 1 | A | queued |
| cell borders (new layer, second run, restroke, cleared edges) | 2 | A | queued |
| body shape styles: set (six values), rewrite (two), reset (five), reads | 4 | set B (fresh registry UUID); rewrite, reset, reads A | queued |
| body line endpoints: set, reset and reads | 2 | set B (fresh registry UUID); reset and reads A | queued |
| builder language and locale with body text (fr, ja, de-CH; empty and multi-paragraph bodies) | 1 | B (the registry's object UUIDs by form; the leaf build takes the host output's identity) | queued |
| table add (3×4 at offset 0) | 1 | B (the table UUID, and the owner UUID and owner-map entry derived from it, by form; fresh metadata UUIDs) | queued |
| media remove (image native + host-built; movie; audio) | 3 | image B (ZIP layout), movie and audio A | queued |
| media duplicate (image native + host-built; movie; audio) | 3 | B (fresh UUIDs; Document.iwa byte-identical) | queued |
| media replace (image data, movie data and poster, audio data, asset) and asset reads | 5 + read sweep over 6 documents | movie/poster/audio A, image and asset B (ZIP layout), reads A | queued |
| media add (image, movie, audio; native and blank documents) | 3 | B (fresh UUIDs; new members' ZIP layout) | queued |
| populated body-table duplicate (three focused cases, merge-owner refusal, plus corpus sweep) | 4 focused + 17 corpus routes; 1 merged-table route refused before publish | B (fresh UUIDs and formula-owner identifiers by policy; untouched members exact; merged geometry is a terminal typed refusal pending shared owner rebinding) | queued |

Tier B for creations: both routes draw a random `text_attribute_uuid_string`
for a new smart field, so the case registers a known-field decoder that
compares that UUID by form (`bookmark_without_uuid`, `date_time_without_uuid`
in `pages_text_cases.rs`); every other byte of the new object and every
untouched object is compared as usual.

Inputs: native fixtures where they exist (`image-adjustments-native`,
`audio-playback-native`, `basic`, `body-sections-unicode`,
`body-table-catalog-native`, `body-table-cells-native`,
`body-table-comment-hidden-native-saved`, `body-table-merges-native`); movies,
host-authored audio, shapes, bookmarks, fields and attachments come from the
host's own constructors (`Draft::create_with_text` + `add_body_movie`,
`add_body_audio`, `add_body_rectangle`, `add_body_bookmark`, …), because the
corpus has no native file with them.

### Pinned divergences (routes left host-only)

| Route | Divergence | Owner |
|---|---|---|
| `duplicate_table` formula-owner allocation | the Pages host repeats an owner-map internal identifier when the source map already contains the next owner value; the leaf allocates above the complete map. Tier B compares the cloned owner family by form and preserves exact reads; volatile, spanning, whole-owner, error and spill dependencies are refused on both routes | W1-B / policy 0368a |
| `duplicate_table` merged-cell owner | the shared structural clone currently copies a non-empty inline `TST.TableModelArchive.merge_owner` without rebinding its owner UUID to the new table. Pages preflights merge geometry and returns typed `Refusal::UnsupportedDependency` before allocation/publication, so it cannot publish an unreadable copied merge graph; the focused regression records the host acceptance and leaf safety refusal. Rebinding remains a shared-engine/W1-C follow-up | W1-B / shared duplicate engine |

The `add_table` route was pinned here until W1-A aligned the shared table graph
(record 0370f); it is now migrated (tier B). The duplicate route's four
dependency refusals are intentional and terminal; they are recorded as parity
cases rather than silently falling back to the host.

The number-attachment removal divergence pinned in the first slice was a
host defect and is fixed (§3, item 3).

## 3. Host defects found

1. **Tile row count on cell writes** — the host's cell writer recomputed
   `TST.Tile.numrows` as the row extent; native saves count the stored row
   records (record 0370). Fixed in the host (`update_tile`).
2. **Column-header cell counts** — the host kept only the row-header count;
   native files keep both (`body-table-cells-native.pages`: columns 5, 5, 3,
   3 beside rows 4, 4, 4, 2, 2). Fixed in the host through the shared leaf
   writer `table_write::header_adjust`; a source that never counted the
   column keeps its bucket. Two host tests that pinned the old counts were
   updated. Both changes are queued for macOS on the three set-cell routes.
3. **Number-attachment removal kept an orphan object.** The host deleted
   the U+FFFC and its table entry, but its text-rewrite pruning proof fails
   when the storage header lists references the payload does not count
   (application-saved storages do), so the attachment object, the header
   reference to it and its metadata stayed. Fixed for Pages storages
   (`text::number_attachment::prune_exclusive_number_attachment`, reached
   through `text::Draft::remove_text_number_attachment_pruning` from the
   three Pages remove routes): after the existing exclusivity proof the
   header reference is dropped and the object deleted through the existing
   `remove_unreferenced_number_attachment_objects` (external references, UUID,
   watermark suffix). The shared Keynote/Numbers text route is unchanged
   (W1-D's leaf matches it). Body and first-header removal are now tier A
   (`pages_body_number_attachment_remove_parity`,
   `pages_header_number_attachment_remove_parity`); queued for macOS.

4. **Media removal failed on native files.** Removing a body image,
   movie or audio clip deletes the U+FFFC through the shared text rewrite,
   which leaves the attachment in the body storage header when the header
   lists references the payload does not count (the item 3 cause);
   the host's own "remains referenced" check then refused the removal
   (`image-adjustments-native.pages`). Fixed in the host
   (`pages::editor::prune_removed_body_attachment`, called by the three
   remove routes); queued for macOS.

## 4. Performance

Callgrind (instructions, `Ir`), release build with debug info, CPU 3
pinned, each probe running host, leaf, leaf, host (ABBA) in
`crates/litchi-iwa/tests/parity/pages_cost_probes.rs`; one route = open,
one edit, commit and write. Numbers are the sum of the two runs of each side.

| Route (fixture) | Host | Leaf (before reopen) | Leaf (final) | Δ final |
|---|---:|---:|---:|---:|
| body image vertical flip (`image-adjustments-native`) | 226,888,165 | 74,464,369 | 50,615,218 | −77.7% |
| body image properties | 227,133,348 | 74,476,924 | — | −67.2% (before reopen) |
| body movie geometry (source-built) | 15,394,217 | 4,460,828 | — | −71.0% (before reopen) |
| body audio playback (`audio-playback-native`) | 224,345,681 | 65,368,057 | — | −70.9% (before reopen) |
| shape text set (source-built) | 20,008,506 | 4,415,859 | — | −77.9% (before reopen) |
| whole-body set (`basic`) | 63,771,909 | 69,600,252 (+9.1%) → 65,628,626 (+2.9%) after `234929c79b` | 41,843,055 | −34.4% |
| bookmark add (`basic`) | 114,216,883 | 86,511,501 | 63,189,426 | −44.7% |
| Date & Time insert (`basic`) | 129,487,304 | 86,834,529 | — | −32.9% (before reopen) |
| number attachment insert (`basic`) | 120,964,310 | 86,901,649 | — | −28.2% (before reopen) |
| table cell value (`body-table-catalog-native`) | 2,532,811,983 | 78,390,305 | 78,390,305 | −96.9% |
| table cell values, 16-cell batch | 2,788,732,498 | 110,195,326 | — | −96.1% (before reopen) |
| number attachment remove (`basic` + host page count) | 159,628,677 | — | 52,679,103 | −67.0% |
| image title create (`image-adjustments-native`) | 546,895,891 | — | 85,626,043 | −84.3% |
| image caption remove (host-labelled image) | 710,021,022 | — | 85,975,944 | −87.9% |
| section append (`body-sections-unicode`) | 589,774,024 | — | 127,211,178 | −78.4% |
| section remove | 498,795,655 | — | 137,862,844 | −72.4% |
| section insert (split, `body-sections-unicode`) | 592,827,466 | — | 128,252,805 | −78.4% |
| body-table remove (`body-table-catalog-native`) | 2,246,224,218 | — | 170,815,958 | −92.4% |
| image remove (`image-adjustments-native`) | 659,210,379 | — | 66,936,803 | −89.9% |
| movie remove (source-built) | 36,674,391 | — | 5,525,855 | −84.9% |
| image duplicate (`image-adjustments-native`) | 511,702,845 | — | 91,585,043 | −82.1% |
| image data replace (`image-adjustments-native`) | 234,144,941 | — | 51,627,431 | −78.0% |
| media assets read (open + list) | 55,302,013 | — | 45,294,193 | −18.1% |
| image add (`image-adjustments-native`) | 428,092,047 | — | 79,769,917 | −81.4% |
| table sort apply (`body-table-catalog-native` table 0) | 3,397,895,357 | — | 54,905,570 | −98.4% |
| table resize (`body-table-catalog-native` table 0 to 7×6) | 2,531,579,334 | — | 50,760,582 | −98.0% |
| table resize, after the owner-extents fix (both sides) | 2,541,546,946 | — | 51,157,126 | −98.0% |
| table add (`body-table-catalog-native`, 3×4 at offset 0) | 3,262,277,019 | — | 147,383,467 | −95.5% |
| table add, after W1-A's one-pass UID-map encoder (trunk `e658e4676e`; parity re-run, tier B) | 3,262,401,911 | — | 147,276,483 | −95.5% |
| cell fill set (table 0, cell 1,1) | 3,555,552,959 | — | 130,637,922 | −96.3% |
| cell layout reset (host-styled cell 1,1) | 2,813,295,373 | — | 209,686,265 | −92.5% |
| cell border set (table 0, top edge of cell 1,1) | 3,381,823,315 | — | 89,096,263 | −97.4% |
| body shape fill set (host rectangle on `chart-data-native`) | 416,949,768 | — | 142,213,368 | −65.9% |
| builder: body `Bonjour`, `fr`/`fr_FR` | 3,565,170 | 7,061,105 (four publications) → 5,489,195 (two) | 3,499,471 (one) | −1.8% |
| body line endpoints set (`chart-data-native`) | 423,638,442 | — | 143,606,634 | −66.1% |
| body line endpoints reset (`chart-data-native`) | 549,682,634 | — | 127,418,970 | −76.8% |
| body line endpoints read (`chart-data-native`) | 155,394,765 | — | 31,411,145 | −79.8% |
| populated body-table duplicate (`body-table-cells-native`) | 4,215,551,178 | — | 138,534,970 | −96.7% |

### 4.1 Resumed qualification on the final work-package candidate

The boundedness and merge-owner repairs were remeasured after merging the
current coordinator trunk. The source was `6bb1d58021` (branch
`uo2/iwa-w1b`, coordinator parent `0944b7319f`). The release parity binary was
built from that exact source with
`CARGO_PROFILE_RELEASE_DEBUG=1 CARGO_BUILD_JOBS=4 cargo test --release -p
litchi-iwa --all-features --test parity --no-run`. The prescribed probe then
used CPU 3, callgrind instruction counts, and the existing host/leaf ABBA
probe pair:

```sh
bash /home/zhuhe/code/litchi-worktrees/.coord/iwa-w1b/cg.sh \
  target/release/deps/parity-bf69cf43b955a09c \
  pages_shape_style_cases::line_endpoints_set_cost_abba \
  pages_shape_style_cases::line_endpoints_reset_cost_abba \
  pages_shape_style_cases::line_endpoints_read_cost_abba \
  pages_lifecycle_cases::table_duplicate_cost_abba
```

| Final-candidate route | Host Ir | Leaf Ir | Leaf delta |
|---|---:|---:|---:|
| line endpoints set | 423,509,798 | 143,612,630 | −66.09% |
| line endpoints reset | 549,099,195 | 127,434,349 | −76.79% |
| line endpoints read | 155,392,745 | 31,413,766 | −79.78% |
| supported body-table duplicate | 4,220,277,802 | 139,246,763 | −96.70% |

All supported routes remain below the host cost; the no-more-than-5%/20 µs
bar is met. The merged-cell route is intentionally refused before candidate
publication, so it has no mutation cost to qualify. These are full parity
facade probes (not engine-only measurements).

The raw callgrind outputs and binary remain outside the repository. This
manifest makes the receipt reproducible without committing large evidence:

| Artifact | Size | SHA-256 |
|---|---:|---|
| `/home/zhuhe/code/litchi-worktrees/uo2-iwa-w1b/target/release/deps/parity-bf69cf43b955a09c` | 198,830,136 | `66ede694ec732bb3150a34823d649a77e22dfddcee224094998ce9167ce15056` |
| `/home/zhuhe/code/litchi-worktrees/.coord/iwa-w1b/cg_results.txt` | 3,298 | `0e805233b8b6f5372ebe01ecd348c2186b72da3394b5c3a643e0e711f6418daf` |
| `cg/pages_shape_style_cases::line_endpoints_set_cost_abba.host.out` | 592,436 | `80cef034853a871d9e80caa9f86a3960fe4d5e52ac1a3b21668b6d04cbd04e0b` |
| `cg/pages_shape_style_cases::line_endpoints_set_cost_abba.leaf.out` | 480,629 | `f5b0aa6c4b6e4d6a33ca0aca2f5ed47132ab32ff6620960bc73dbfe72b69facf` |
| `cg/pages_shape_style_cases::line_endpoints_reset_cost_abba.host.out` | 565,209 | `f851d2587642a30ebe9d1325aa5c71ad5c6932937ace5c8aa68bf6547fc7eec4` |
| `cg/pages_shape_style_cases::line_endpoints_reset_cost_abba.leaf.out` | 449,489 | `05460e47157c3cedb9c5d757f0e01f16ca263ba4efbe0bcb3e3e236dea280244` |
| `cg/pages_shape_style_cases::line_endpoints_read_cost_abba.host.out` | 382,185 | `61cbfdc3c3c97201a6b585f151a58924be12058cd156ebd4a9ee0ee706dc40d6` |
| `cg/pages_shape_style_cases::line_endpoints_read_cost_abba.leaf.out` | 297,556 | `298db18e0dec06a098f37d1ac0609643e4f4c42f09126ffdd6368df30804192e` |
| `cg/pages_lifecycle_cases::table_duplicate_cost_abba.host.out` | 830,838 | `5ec3f079e18882862d2e97e78fc63c3795dfd3be0ad54e8ee8efebc5b8283cd` |
| `cg/pages_lifecycle_cases::table_duplicate_cost_abba.leaf.out` | 780,391 | `b78d89938a5e8c7e42a3c91a75323be088b859bcb950e32688c89771560dfc8a` |

The raw files are relative to `/home/zhuhe/code/litchi-worktrees/.coord/iwa-w1b`.

The whole-body set route was the only miss: the leaf `edit_body_text` (an
existing family) re-encoded every archive header in its physical preflight
and reparsed the component it edits. `234929c79b` removes both (a unit test
pins the summed length to the encoder on three native fixtures); then the
shared reopen removed the candidate's reparse of untouched members. Probe
commands:

```sh
CARGO_PROFILE_RELEASE_DEBUG=1 cargo test --release -p litchi-iwa --all-features --test parity --no-run
taskset -c 3 valgrind --tool=callgrind '--toggle-collect=*probe_host*' \
  target/release/deps/parity-<hash> --ignored --exact <case>::<probe> --test-threads=1
# and '--toggle-collect=*probe_leaf*'
```

## 5. Verification (this slice)

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-pages -p litchi-iwa -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings`, and `--features pages` | 0, 0 |
| `cargo test -p litchi-pages --all-features` (729) / `-p litchi-iwa --all-features` (1,963, parity target 169 + 15 ignored probes) / `-p litchi --all-features` (209) | 0 / 0 / 0 |
| resumed boundedness follow-up: `cargo test -p litchi-pages --all-features`; `cargo test -p litchi-iwa-style --all-features` (7 + doctests); focused parity `pages_body_line_endpoints` (2), `pages_body_table_duplicate_merge_owner_readback` (1), and `pages_body_table_duplicate_corpus_parity` (17 routes + 5 refusals) | 0 / 0 / 0 |
| resumed `cargo clippy -p litchi-pages -p litchi-iwa-style -p litchi-iwa --all-targets --all-features -- -D warnings` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-pages` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile pages` | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_iwa_exit_ledger` (1,289) | 0 |
| `python3 tools/{officecli_capability_delta,officecli_corpus,unified_bench_matrix,iwa_exit_ledger}.py --check`; `check_crate_boundaries.py`; `non_iwork_gate.py verify`; `check_feature_matrix.py` | 0 each |
| `cargo test -p litchi-ops --all-features` in chunks: `--lib` (1,022), `--doc` and nine groups of up to 20 `--test` targets (1,702) | 0 each |
| `crates/pyo3-litchi/tests/run_tests.sh` (301) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |

Slices 5–6 (section split, body storage, sort, resize, table add, the
owner-extents and body-owner fixes) ran the same commands on `01ebb98c65`
(merged with trunk `dec1a5a4e2`): every exit 0 except
`check_crate_boundaries.py` (1: the section-text owner exposed `Storage`
and kept the flat `body_text`), fixed in `f44c4ea905` and re-run (0, with
the full glossary gate 0) after merging trunk `32f2c3ad77`.

No litchi-ops code, capability or corpus-visible behaviour changed, so the
capability/registry goldens and the OfficeCLI corpus were not re-blessed.
Both pre-commit size checks printed nothing for every commit.

The resumed hardening keeps the same public route contract while closing three
boundedness/safety holes found during coordinator review: body line-end field
scans now use the owning `Budget` (including the flip varint and retained
endpoint copies), and the shared line-end codec has one bounded
`decode(payload, limits)` entry point whose structured error keeps semantic
unknown-endpoint, wire-limit and allocation failures distinct. The Pages
facade maps those wire failures to its typed `Limit`/`Allocation` errors;
malformed endpoint semantics remain the terminal `UnsupportedStyle` refusal.
Duplicate-table candidate readback now preserves `table_lock` limit and
allocation failures instead of turning them into `Verification`. Finally,
duplicate preflight reads the selected table's merge geometry under its own
bounded reader, preserves merge-reader limit/allocation distinctions, and
returns `UnsupportedDependency` for non-empty merge owners until the shared
clone engine can rebind them.

## 6. Gaps and follow-ups (rest of W1-B)

| Gap | Rows | Next step / dependency |
|---|---:|---|
| text-box attachments (incl. removal) and date-time fields | 5 + 5 | text boxes are W2-B |
| shape lifecycle; shape fill, stroke, shadow, effects, text layout and preset | ~34 | W1-E codecs (`shape_style_codec` straight lines landed); line endpoints are migrated in this slice |
| table cell formats, rich text, topology, comment writes, conditional highlights | ~120 | cores in `litchi-iwa-table` (W1-G, record 0379); W1-B adds the Pages facades when each group lands |
| image cell fills | — | data-reference registry bookkeeping for a style variation (host `register_style_object` with a data identifier) |
| table formulas | — | W1-F formula engine (record 0376) |
| Builder `body_table` | 1 | the host bootstraps the first table graph from its Numbers builder; the leaf `table::creation` needs a template table, so a table-less document needs that bootstrap first |
| non-empty merged-cell table duplication | 1 focused/corpus route | Pages refuses before publication because the shared duplicate engine does not yet rebind the inline merge-owner UUID; W1-C/shared duplicate-engine work must repair the owner remap before this route can move from the explicit safety refusal to parity |
