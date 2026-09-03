# IWA-protos fuzz manifest

`Cargo.toml` is the source of truth for runnable targets. The table below
keeps the checked-in target/corpus relationship reviewable without making
generated libFuzzer output part of the repository.

| Registered target | Harness | Seed directory |
| --- | --- | --- |
| `numbers_tile_storage` | `fuzz_targets/numbers_tile_storage.rs` | `corpus/numbers_tile_storage/` |
| `numbers_formula_archive` | `fuzz_targets/numbers_formula_archive.rs` | `corpus/numbers_formula_archive/` |
| `pages_section_codec` | `fuzz_targets/pages_section_codec.rs` | `corpus/pages_section_codec/` |
| `pages_section_text_codec` | `fuzz_targets/pages_section_text_codec.rs` | `corpus/pages_section_text_codec/` |
| `pages_body_footnote_codec` | `fuzz_targets/pages_body_footnote_codec.rs` | `corpus/pages_body_footnote_codec/` |
| `pages_footnote_graph_codec` | `fuzz_targets/pages_footnote_graph_codec.rs` | `corpus/pages_footnote_graph_codec/` |
| `pages_footnote_codec` | `fuzz_targets/pages_footnote_codec.rs` | `corpus/pages_footnote_codec/` |
| `pages_movie_caption_codec` | `fuzz_targets/pages_movie_caption_codec.rs` | `corpus/pages_movie_caption_codec/` |
| `pages_header_footer_codec` | `fuzz_targets/pages_header_footer_codec.rs` | `corpus/pages_header_footer_codec/` |
| `numbers_table_data_list` | `fuzz_targets/numbers_table_data_list.rs` | `corpus/numbers_table_data_list/` |
| `table_dimension` | `fuzz_targets/table_dimension.rs` | `corpus/table_dimension/` |
| `table_info_codec` | `fuzz_targets/table_info_codec.rs` | `corpus/table_info_codec/` |
| `table_appearance` | `fuzz_targets/table_appearance.rs` | `corpus/table_appearance/` |
| `numbers_table_sort_order_codec` | `fuzz_targets/numbers_table_sort_order_codec.rs` | `corpus/numbers_table_sort_order_codec/` |
| `numbers_table_header_settings_codec` | `fuzz_targets/numbers_table_header_settings_codec.rs` | `corpus/numbers_table_header_settings_codec/` |
| `numbers_table_title_codec` | `fuzz_targets/numbers_table_title_codec.rs` | `corpus/numbers_table_title_codec/` |
| `table_model_discovery_codec` | `fuzz_targets/table_model_discovery_codec.rs` | `corpus/table_model_discovery_codec/` |
| `comment_storage_codec` | `fuzz_targets/comment_storage_codec.rs` | `corpus/comment_storage_codec/` |
| `comment_storage_reply_codec` | `fuzz_targets/comment_storage_reply_codec.rs` | `corpus/comment_storage_reply_codec/` |
| `keynote_chart_title` | `fuzz_targets/keynote_chart_title.rs` | `corpus/keynote_chart_title/` |
| `keynote_chart_legend_codec` | `fuzz_targets/keynote_chart_legend_codec.rs` | `corpus/keynote_chart_legend_codec/` |
| `keynote_chart_axis_title_codec` | `fuzz_targets/keynote_chart_axis_title_codec.rs` | `corpus/keynote_chart_axis_title/` |
| `keynote_chart_axis_value_settings_codec` | `fuzz_targets/keynote_chart_axis_value_settings_codec.rs` | `corpus/keynote_chart_axis_value_settings_codec/` |
| `keynote_chart_caption` | `fuzz_targets/keynote_chart_caption.rs` | `corpus/keynote_chart_caption/` |
| `keynote_movie_caption` | `fuzz_targets/keynote_movie_caption.rs` | `corpus/keynote_movie_caption/` |
| `keynote_movie_geometry_codec` | `fuzz_targets/keynote_movie_geometry_codec.rs` | `corpus/keynote_movie_geometry_codec/` |
| `numbers_table_cell_pop_up_menu_codec` | `fuzz_targets/numbers_table_cell_pop_up_menu_codec.rs` | `corpus/numbers_table_cell_pop_up_menu_codec/` |
| `numbers_table_cell_control_codec` | `fuzz_targets/numbers_table_cell_control_codec.rs` | `corpus/numbers_table_cell_control_codec/` |
| `numbers_table_cell_number_format_codec` | `fuzz_targets/numbers_table_cell_number_format_codec.rs` | `corpus/numbers_table_cell_number_format_codec/` |
| `numbers_table_cell_percentage_format_codec` | `fuzz_targets/numbers_table_cell_percentage_format_codec.rs` | `corpus/numbers_table_cell_percentage_format_codec/` |
| `numbers_table_cell_currency_format_codec` | `fuzz_targets/numbers_table_cell_currency_format_codec.rs` | `corpus/numbers_table_cell_currency_format_codec/` |
| `numbers_table_cell_scientific_format_codec` | `fuzz_targets/numbers_table_cell_scientific_format_codec.rs` | `corpus/numbers_table_cell_scientific_format_codec/` |
| `numbers_table_cell_fraction_format_codec` | `fuzz_targets/numbers_table_cell_fraction_format_codec.rs` | `corpus/numbers_table_cell_fraction_format_codec/` |
| `movie_playback_codec` | `fuzz_targets/movie_playback_codec.rs` | `corpus/movie_playback_codec/` |
| `keynote_table_physical_sort_codec` | `fuzz_targets/keynote_table_physical_sort_codec.rs` | `corpus/keynote_table_physical_sort_codec/` |
| `package_metadata_media_codec` | `fuzz_targets/package_metadata_media_codec.rs` | `corpus/package_metadata_media_codec/` |
| `pages_drawable_order_codec` | `fuzz_targets/pages_drawable_order_codec.rs` | `corpus/pages_drawable_order_codec/` |

`fuzz_targets/numbers_table_model.rs` is retained as a source-only harness and
is intentionally not registered until it is revalidated. A source file alone
is not a runnable cargo-fuzz target. Corpus recipes for that source-only
harness remain review material under `corpus/numbers_table_model/`; they are
not consumed by `cargo fuzz run`.

The checked-in recipes are hand-authored `hex:` inputs where documented by the
target README. Empty seed directories are valid: those targets build bounded
valid and malformed inputs in the harness or exercise arbitrary bytes directly.
Do not copy generated corpus additions or artifacts into this tree.

To audit the manifest without compiling the workspace:

```sh
cargo metadata --no-deps --format-version 1
cargo +nightly fuzz list
```

For a bounded local check, keep mutable corpus, artifacts, and build output
outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-iwa-protos-fuzz.XXXXXX")"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz check \
  --fuzz-dir "$PWD" --target-dir "$fuzz_root/target" numbers_table_data_list
```
