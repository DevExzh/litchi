# 0371 — W1-B: Pages body media, shapes, text fields and table cells out of the host

- Status: first merge-ready slice on branch `uo2/iwa-w1b` (cut from trunk
  `feat/unified-ops-api` at `f35f1c9f55`; trunk merged several times, last
  at the head of this slice). The wave continues (§6).
- Date: 2026-10-26.
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
| `2365445db4` | `Package::{body_storage, body_text}`: the rooted body storage (`litchi_iwa_text::storage::Storage`) and its complete native text, replacing the host reads of the same names | 2 |
| `b5af407b34` | `table::sort::execution` — `Package::{edit_body_table_sort_execution, apply_body_table_sort_execution}`; `Edit::{apply, apply_to_rows}`: the host's scalar executor at wire level (stable permutation by the shared planner; tile row records, row header buckets, row UID map and stroke layers move; user-hidden rows restored to their positions in the active hidden-state extent). The payload rewrites take only payloads and a permutation, for the shared table crate | 2 |
| (table resize commit, below) | `table::size` — `Package::{edit_body_table_size, apply_body_table_size}`; `Edit::set(rows, columns)`: the host's resize (tiles trimmed, header buckets kept inside the extent and rewritten with explicit cell counts, UID-map identities appended or dropped, stroke-sidecar and model counts) | 1 |
| `acf5fd71ad` | crate-boundary gate: the bookmark, date-time and attachment re-export modules move to their own files; the movie-playback owner keeps its feature-gated declaration with a `#[path]` twin for builds without the host seam; the retired-host-route audit skips the parity harness (whose cases call host routes by design) | — |
| `234929c79b`, `28fb68c468` | performance (§4): the section-text preflight sums retained header lengths instead of re-encoding the archive and clones the parsed component instead of reparsing it; every Pages commit-time candidate reopen calls `SourceCatalog::reopen` | — |

A private engine `package::body_storage_objects` holds what creating and
deleting a storage-owned object needs (the host's rules): the identifier one
above every physical object and metadata registration, the metadata
watermark raise, external-reference and UUID-registration removal and the
watermark-suffix release, and the TSWP text splice.

In all, 60 Pages ledger rows flip to `migrated` in this slice (63 with
W0-C's merge rows). Every new public type follows ADR 0034; the families are
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
| media remove (image native + host-built; movie; audio) | 3 | image B (ZIP layout), movie and audio A | queued |
| media duplicate (image native + host-built; movie; audio) | 3 | B (fresh UUIDs; Document.iwa byte-identical) | queued |
| media replace (image data, movie data and poster, audio data, asset) and asset reads | 5 + read sweep over 6 documents | movie/poster/audio A, image and asset B (ZIP layout), reads A | queued |
| media add (image, movie, audio; native and blank documents) | 3 | B (fresh UUIDs; new members' ZIP layout) | queued |

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
| `add_table` | the leaf `table::creation` family (record 0329) writes a different graph from the host (random table UUID, its own identifier order, formula owner type 6267 against 4008) — `pages_body_table_add_divergence_is_reported` | W1-B follow-up: align the leaf creation with the host graph, or document the leaf graph as the native one and retire the host |

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
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-pages` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile pages` | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_iwa_exit_ledger` (1,289) | 0 |
| `python3 tools/{officecli_capability_delta,officecli_corpus,unified_bench_matrix,iwa_exit_ledger}.py --check`; `check_crate_boundaries.py`; `non_iwork_gate.py verify`; `check_feature_matrix.py` | 0 each |
| `cargo test -p litchi-ops --all-features` in chunks: `--lib` (1,022), `--doc` and nine groups of up to 20 `--test` targets (1,702) | 0 each |
| `crates/pyo3-litchi/tests/run_tests.sh` (301) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |

No litchi-ops code, capability or corpus-visible behaviour changed, so the
capability/registry goldens and the OfficeCLI corpus were not re-blessed.
Both pre-commit size checks printed nothing for every commit.

## 6. Gaps and follow-ups (rest of W1-B)

| Gap | Rows | Next step / dependency |
|---|---:|---|
| text-box attachments (incl. removal) and date-time fields | 5 + 5 | text boxes are W2-B |
| `add_table` | 1 | the leaf table creation (record 0329) writes a different graph from the host's `add_table` (random table UUID, its own identifier order, formula-owner type 6267 against the host's 4008), so both stay host-only until one side is aligned |
| shape lifecycle; shape fill, stroke, shadow, effects, text layout, line ends, preset | ~37 | W1-E codecs (`shape_style_codec` straight lines landed) |
| table cell formats, rich text, appearance, conditional highlights, formulas, topology, duplicate/resize, sort execution, comment writes | ~150 | W1-A engines (topology engine landed; formats/styles follow), W1-F formula caches, W1-D paragraph/list engines |
| Builder `language`, `locale`, `body_table` | 6 | next (W1-C did the Keynote pair) |
