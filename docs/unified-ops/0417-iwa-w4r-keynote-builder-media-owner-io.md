# 0417 — W4-R: Keynote builder, media-asset replacement, host owner I/O and leftovers

- Status: landed on branch `uo2/iwa-w4r` (cut from trunk `12682616f7`,
  trunk `8010cc9fb4` merged). Ready for the coordinator merge once the
  gate list in §6 is green.
- Date: 2026-11-02.
- Scope (coordinator): the remaining non-table, non-chart host-only rows of
  `results/0363/iwa-exit-ledger.tsv`: (1) the Keynote builder, (2) Keynote
  `replace_media` and the shared `media/package.rs` rows, (3) small
  leftovers (`document.rs` media/text/application, Numbers
  `sheet_text_boxes`, `topology_oracle`, `lib.rs` publication accessors,
  `package.rs` raw rows and limit constructors), (4) host owner I/O rows.
- `performance_claim`: the Keynote builder and the Keynote asset
  replacement are cheaper than the host routes they replace (§4).

## 1. What landed

| Commit | What |
|---|---|
| `8b73bc0e01` | `litchi_keynote::Builder::slide_number_visible(bool)` (the host builder's option, through the slide placeholder-visibility family); `Builder::slide_size` also lays out the title and body placeholders of the layout and the first slide for the size, with the host's fractions (title at 10 %/24 %, 80 % × 20 %; body at 15 %/52 %, 70 % × 14 %) and its rectangle path. Before, a sized leaf build kept the blank presentation's 1920×1080 frames: a capability gap against the host |
| `c1c0b03dc9` | `litchi_keynote::slide::media::asset_data` — `Package::{edit_media_asset_data, edit_media_asset_data_with_options, apply_media_asset_data}`, `Edit::{asset, before, after, set, commit}`, `EditOptions::max_replacement_bytes(NonZeroUsize)` (1 GiB default, the host's), `Patch::{is_noop, inverse}`, `Diagnostics::{changed, deleted_previews}`. Replaces one asset that `Package::media_assets` lists (image, movie content or poster, audio) keeping its identifier; the `Data/` member and its `DataInfo` digest and length are rewritten by the existing slide-media owner (`rewrite_member`, factored out of `rewrite_media`); root previews are invalidated; the candidate is checked for read-back and ZIP locality. New refusals `AssetNotReachable`, `AssetNotMaterialized`. Registered in the `keynote-minimal`/`keynote-full` glossary profiles |
| `213a5d31af` | parity: Keynote builder (six cases) and asset replacement (three cases); 15 ledger rows |
| `f8bd398e5e` | parity: host owner I/O, `Document` reads, `sheet_text_boxes`, Pages `create`/`create_with_text`; 71 ledger rows |
| `583901cb2a` | performance: the Keynote builder writes language, locale, slide size (with the placeholder layout), title, subtitle and notes in one publication (`Package::with_fresh_settings`); `builder::Error::{Locale, ShowSettings}` become `builder::Error::Settings` |

## 2. Ledger

86 rows change: 17 migrated, 69 retired. Host-only rows after the merge:
Keynote 50, Numbers 101, Pages 64, shared 1 (`document.rs charts`, chart
lifecycle, W4-D). None of the rows in this package's scope stays
host-only.

### 2.1 Migrated

| Host rows | Leaf route | Parity | Tier |
|---|---|---|---|
| `keynote/creation.rs` `new`, `builder`, `build`, `create` | `litchi_keynote::Builder::{new, build}` (`create` = a default build with `Identity::generate()`) | `keynote_builder_defaults_parity` | B |
| `title`, `subtitle`, `presenter_notes` | `Builder::{title, subtitle, notes}` | `keynote_builder_text_parity` (+ `keynote_builder_empty_text_parity`) | B |
| `language`, `locale` | `Builder::{language, locale}` | `keynote_builder_language_locale_parity` (fr, ja, de-CH, alone and with text; language or locale alone) | B |
| `slide_size` | `Builder::slide_size` | `keynote_builder_slide_size_parity` (three sizes; one with every option) | B |
| `keynote/editor.rs replace_media` | `Package::edit_media_asset_data` | `keynote_media_asset_replace_parity` (every replaceable asset of three native presentations) | B |
| `media/package.rs replace` | same | `media_draft_replace_parity` (host `media::Draft::replace`) | B |
| `media/package.rs set_max_replacement_bytes` | `slide::media::asset_data::EditOptions::max_replacement_bytes` | `media_asset_replacement_limit_parity` | B |
| `pages/editor.rs create`, `create_with_text` | `litchi_pages::builder::Builder::{build, body_text}` | `pages_builder_create_parity` | B |
| `media/package.rs remove_unreferenced` | one shared rule, `litchi_iwa_style::media::unreferenced_data` (no component owner, object data reference or `DataMetadataMap` entry), with facades `prune_unreferenced` on Keynote `slide::media::asset_data::Edit` (from the new `Package::edit_media_assets`), Pages `drawable::lifecycle::Edit` and Numbers `sheet::drawable::media::Edit` (through the shared data-owner transition's `data_removals`, and its new `data_record_removals` for an unmaterialized record); the leaf prunes every orphan in one commit, the host one identifier at a time | `media_prune_unreferenced_parity` (three formats; one and three orphans; inverse), `media_prune_unmaterialized_parity` | B |
| `numbers/editor/semantic/text_box.rs sheet_text_boxes` | `litchi_numbers::Package::sheet_drawable_list` (`Kind::TextBox`) with `sheet_drawable_plain_text` | `numbers_sheet_text_boxes_read_parity` | A |

`Builder::slide_number_visible` has no ledger row (the host method is a
`const fn`, outside the crosswalk scan); `keynote_builder_slide_number_parity`
covers it.

### 2.2 Retired (each reason names the leaf API and a test pointer)

| Host rows | Leaf replacement | Test pointer |
|---|---|---|
| `media/package.rs insert_unreferenced` | host-internal step of the host media add routes. No caller needs an unattached record: no application shows it, no supported API can attach it later (every leaf add takes the bytes and allocates the record in the same transaction: Pages `edit_body_media`, Numbers `sheet::drawable::Edit::add_*`, Keynote `slide::drawable::Edit::add_image` and the movie/audio creation families), and the leaves prune it as garbage | the lifecycle parity cases of records 0371, 0372, 0378 |
| `comments.rs`, `text/editor.rs`, `pages/editor.rs`, `numbers/editor/semantic/{workbook,media}.rs` owner I/O (`open`, `from_bytes`, `save`, `to_bytes`, `from_package`, `package`, `into_package`) | the immutable leaf `Package::{open, from_bytes, write_to, save}` (Pages also `to_bytes`); the host raw-package conversions go with the host (0363 §3.1) | `pages_package_lifecycle_parity`, `numbers_package_lifecycle_parity`, `shared_draft_lifecycle_parity` |
| `package.rs` raw `Package` (`new`, `open`, `from_bytes`, `from_shared_bytes`, `from_read_at`, `len`, `is_empty`, `validate`, `to_bytes`, `write_to`, `save`, `contains_entry`) | `litchi_iwa_archive::SourceCatalog::{from_bytes, from_shared_bytes, from_read_at}`, `package::Catalog::{len, is_empty, to_bytes, write_to}`, `unstable::entry`, the leaf `save` | `physical_package_owner_parity`, `*_physical_access_parity` |
| `package.rs` raw `Edit` family (`edit`, `commit`, `apply`, `reset`, `insert_entry`, `remove_entry`, `replace_archive`, `insert_archive_before`, `update_archive`, and `len`/`is_empty`/`validate`/`contains_entry` of the staged edit) | deleted with the host (0363 §3.1, D5: no raw mutation API); physical member changes are `litchi_iwa_archive::package::Catalog::{prepare_reassembly_with_changes, reassemble_with_*}` with `EntryEdit` and `ReassemblyPatch::{apply, inverse}` | — |
| `package.rs` `PackageLimits::{new, from_values, with_input_bytes, with_iwa_stream_bytes, with_archive_limits}` | the type is not exported and no remaining host route takes it; `litchi_iwa_archive::Limits::{new, with_archive_limits}` and each format's `SourceLimits` through `OpenOptions` | — |
| `document.rs` `open`, `from_bytes`, `snapshot`, `validate`, `application`, `media_stats`, `extract_media`, `extract_media_to_writer`, `extract_media_to_file`, `text` | the format leaf after `litchi_iwa_detect`, `litchi_ops::{open, open_bytes}`; `media::Catalog::stats` over the members; `Data/<name>` through `unstable::entry` and the leaf asset reads; the leaf semantic text reads (`litchi_ops::Iwork::text`) | `document_read_parity` |
| `lib.rs publication_error`, `publication_was_committed` | the leaf typed `SaveError` (`Publication`, `was_committed`) | — |
| `numbers/editor/topology_oracle.rs __edit_table_topology_unverified` | `doc(hidden)` differential oracle of a host test; not a supported API (0363 §3.1) | — |

## 3. Parity and documented differences

All new cases are in `crates/litchi-iwa/tests/parity/{keynote_builder_cases,
keynote_media_asset_cases, owner_io_cases, pages_builder_cases}.rs`.

- **Builder identities.** The host builder draws a random document
  identity, template slide identifier, theme UUID, table-template UUID and
  row/column UIDs, and object UUIDs. As in record 0371 for Pages, the leaf
  build takes the host output's document identity, so the properties,
  document identifier and registry revision compare exactly. The template
  slide identifier (`KN.SlideNodeArchive` field 29), the theme UUID
  (`KN.ThemeArchive` field 3) and the table UUID (`TST.TableModelArchive`
  field 1) are masked byte-exactly (every other byte of those objects,
  unknown fields included, is compared); a read checks that both slide
  nodes share one template identifier. The UID map and the registry's
  object UUIDs are compared by form, as for Numbers (0370i).
- **Empty text (documented difference).** An empty title, subtitle or notes
  is written by the leaf in the form every native Keynote fixture in
  `test-data/iwork/keynote` uses for an empty storage (no `text` field;
  45–52 such storages per presentation, none with `[""]`); the host
  builder writes `text: [""]`. `keynote_builder_empty_text_parity` compares
  that one field in its native form. Queued for macOS.
- **Asset replacement.** Tier B: the leaf invalidates the root previews
  (the Keynote media policy) and reassembles the ZIP; `Data/` bytes,
  `DataInfo` and every other object compare. The inverse restores the
  input bytes and `apply_media_asset_data` replays the patch.
- **Document text.** The host `Document::text` dumps every TSWP storage in
  archive order (master placeholder text, U+FFFC attachment characters)
  and misses Numbers cell text; the leaf reads are semantic. Pages is equal;
  for Keynote every leaf line occurs in the host dump.
- **Pruning.** Host-inserted orphans (`media::Draft::insert_unreferenced`)
  on a native file of each format; the host removes each with
  `remove_unreferenced`, the leaf prunes them in one commit. Tier B: the
  host re-encodes `TSP.PackageMetadata` with Prost; Keynote also
  invalidates the root previews (its structural-publication policy). The
  inverse restores the input. An unmaterialized orphan (no `Data/`
  member) loses its `DataInfo` on all three formats
  (`media_prune_unmaterialized_parity`); for Numbers the shared
  `litchi_iwa_table::structure::metadata::Transition` gains the additive
  `data_record_removals` (a `DataInfo` removal without a member), the
  member-removing `data_removals` path unchanged.

macOS queue: 11 Keynote rows (builder values, the empty-text form, asset
replacement) and 3 shared rows (format-neutral replace, the ceiling, pruning) in
`0363-mac-queue.md`; the Pages builder and Numbers drawable-list rows
were already queued.

## 4. Performance

Callgrind (`perf` is blocked), release parity binary, `taskset -c 3`,
ABBA (host, leaf, leaf, host); inclusive Ir of the probe functions,
per call.

| Route | Host | Leaf | Δ |
|---|---:|---:|---:|
| Keynote builder, defaults | 13.18 M | 2.77 M | −79 % |
| Keynote builder, title + subtitle + notes | 13.40 M | 4.82 M | −64 % |
| Keynote builder, every option (also fr/fr_FR, 1024×768, slide number) | 12.97 M | 8.20 M | −37 % |
| Keynote asset replacement (`image-adjustments-native.key`, JPEG) | 128.9 M | 80.8 M | −37 % |

Before `583901cb2a` the every-option build cost 19.0 M (+43 % over the
host): each value was its own transaction with a full reopen. Probes:
`keynote_builder_cases::keynote_builder_cost_abba`
(`KEYNOTE_BUILDER_ROUTE=defaults|text|all`) and
`keynote_media_asset_cases::keynote_media_asset_replace_cost_abba`. The
owner I/O rows retire onto existing leaf owners whose open and write paths
are already measured (records 0296, 0306); no new hot path.

## 5. Gaps and follow-ups

- The Keynote builder keeps the slide-number visibility as its own
  transaction (it also adds the slide-number drawable to the slide's
  drawables and z-order); folding it into the single publication would
  cut about 2.5 M Ir more.
- `document.rs charts` stays host-only (chart lifecycle, W4-D).

## 6. Gates

On `02c1262657` (trunk `8010cc9fb4` merged), build env of the rules file;
every command exited 0.

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-keynote -p litchi-iwa -p litchi-pages -p litchi-numbers -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings`, and with `--features keynote` | 0, 0 |
| `cargo test -p litchi-keynote --all-features` | 0 |
| `cargo test -p litchi-iwa --all-features --test parity`, `--lib` | 0, 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-keynote -p litchi-iwa` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (all profiles; and `--profile keynote-full`) | 0 |
| `cargo test -p litchi-ops --all-features` (`--lib`, `--doc`, `--tests`; capabilities and registry goldens unchanged) | 0 |
| `cargo test -p litchi --all-features` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `python3 tools/check_feature_matrix.py` (31 legs) | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate` (1,276 tests) | 0 |
| `python3 tools/officecli_capability_delta.py --check`, `python3 tools/officecli_corpus.py --check` | 0, 0 |
| `python3 tools/unified_bench_matrix.py --check` (matrix regenerated for 0417) | 0 |
| `python3 tools/check_crate_boundaries.py`, `python3 tools/non_iwork_gate.py verify` | 0, 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |

No dependency changed (no fuzz lock refresh). Target directory 5.1 GB.
