# 0436 — Post-W3 cleanup: remove what existed only for the host

The owner's words (2026-11-05): "Completely remove the crate without considering
the historical technical debits." Record 0429 deleted `litchi-iwa` and left its seams
in the leaves (§5). This package removes them. No leaf byte output changed (§5).

## 1. The feature and the hidden seams

`internal-iwork-source` no longer exists in any crate. `check_crate_boundaries.py` now
rejects the feature and the retired seam names in every crate
(`audit_iwa_retired_host_seams`).

The leaves had 76 `#[doc(hidden)]` attributes in `src`. One remains (§4).

### Became public, because a real non-test caller needs them

| Was | Is | Caller |
|---|---|---|
| Pages `__from_prepared_source_with_semantic_limits` | `Package::from_prepared_source(source, semantic)` | `litchi-ops` `format::iwork` |
| Keynote, Numbers `__from_prepared_source` | `Package::from_prepared_source(source, semantic)` | `litchi-ops` `format::iwork` |
| Keynote `__from_shared_source_with_options` | `Package::from_shared_bytes_with_options` (the name Pages and Numbers already use) | `litchi-ops` `format::iwork`, `iwork_directory_bench` |
| `SourceCatalog::__source_owner` | `SourceCatalog::source_owner` | Numbers (about 80 call sites), `litchi-iwa-table` |
| `OwnedSourceCatalog::{__from_owned_bytes, __from_source_owner}` | `SourceCatalog::{from_owned_bytes, from_source_owner}`; the wrapper type is gone | Numbers |

`PreparedSource` and `SourceCatalog` are physical types of the low-level crates, not
native IDs; ADR 0028 allows those crates that vocabulary, and the leaf `Package` is the
physical API. `tools/check_iwork_leaf_public_api.py` covers the semantic surface, which
did not change.

### Became crate-private

| Was | Is | Why |
|---|---|---|
| Pages `__decode_movie_playback_payload`, `__rewrite_movie_playback_payload`, `MoviePlaybackError` | `pub(super)` without the prefix | the drawable-properties family calls them |
| Keynote `decode_slide_drawable_projection` | `#[cfg(test)]` | only its unit tests call it |
| Keynote `optional_settings_from_projection` | `#[cfg(test)]` | only its unit tests call it |
| Detect `__into_components` | `#[cfg(test)] into_components` | only its frozen-entry tests call it |

### Deleted, with nothing outside using them

Numbers: `SourceBuiltAppearancePayload`, `__table_appearance_from_source_built` and its
budget helpers; `__compatibility_tables_from_prepared_source`;
`__semantic_document_from_prepared_source`; `__decode_image_adjustments_payload`,
`ImageAdjustmentsError`; `__movie_playback_settings`,
`__rewrite_movie_playback_settings`, `MoviePlaybackError` and the `movie_playback` module;
`__from_shared_source_with_options`; `SemanticLimits::__from_archive`;
`MergeReader::__from_shared_catalog`; the modules `table_sort_compat`,
`table_relocation_compat` and `comments_compat`; the `compatibility_defaults` comment mode
and the `preserve_missing_parent` relocation policy, which only those modules set.

Pages: `__CatalogBodyTableAppearanceSource`, `__catalog_body_table_appearance`,
`__catalog_body_table_style_edges`; `__decode_image_adjustments_payload`,
`ImageAdjustmentsError`; `__is_valid_pages_text_storage`, `__pages_body_storage_discovery`,
`BodyStorageDiscovery` and the `text_storage` module; `__from_prepared_source`;
`__semantic_document_from_prepared_source` and the route behind it;
`MergeReader::__from_shared_catalog`.

Keynote: `__CatalogTableAppearanceSource`, `__catalog_table_appearance`,
`__catalog_table_style_edges`; `__decode_slide_drawable_projection`,
`__decode_slide_node_projection`; `__presentation_from_prepared_source`;
`__decode_image_adjustments_payload`, `ImageAdjustmentsError`;
`__invalidate_slide_preview`, `SlidePreviewInvalidationError`;
`__is_litchi_source_built_compatibility`; `__transition_settings_from_source`;
`__slide_table_appearances` and its batch scan mode; `MergeReader::__from_shared_catalog`.

Shared: `ComponentCatalog::__from_shared_archives` and its module; the
`litchi-iwa-detect` profiles `__from_{bytes,shared_bytes,path}_with_pages_metadata`,
`__from_{bytes,shared_bytes,path}_with_semantic_metadata`, `__from_path_with_properties`,
`__into_pages_semantic_source`.

No alias of a removed name exists.

## 2. Tests

Rewritten to pin leaf behaviour through a public route, on native fixtures:

- `native_audio_playback` (Numbers, Pages): the public drawable-properties family instead of
  the payload seam. It asserts more than before: an unchanged `set_playback` now returns the
  source byte for byte.
- `native_image_adjustments` (Numbers): a test-local decoder over the shared codec.
- `prepared_source` (Numbers, Keynote): `Package::from_prepared_source`.
- Pages `package.rs` prepared-source tests: section and text limits are now asserted on
  `Package::from_prepared_source`, the live route.
- `merge_reader_ownership` (Numbers): the object-limit assertion moved to
  `MergeReader::from_bytes_with_options`.
- Pages merge-reader budget tests: moved to unit tests in `body_table_merges/reader.rs`.
- Keynote `missing_name_scan_accumulates_references_across_many_slides`: builds a real package
  instead of parsed archives.
- Keynote `slide_table_appearance`: the batch tests now read each table.
- Detect: three sidecar tests retargeted from the Pages route to the Numbers route.

Deleted:

| Test | Reason |
|---|---|
| `litchi-iwa-archive` `shared_*` (6) | `__from_shared_archives` is gone |
| `litchi-iwa-detect` `pages_zip_profile_*` (7), `pages_shared_and_regular_file_profiles_*`, `directory_pages_*` (2), `pages_metadata_path_profile_*`, `semantic_prepared_source_skips_keynote_media_*` | they test detect profiles no crate calls |
| Numbers `source_built_*` (7) | the source-built bridge is gone |
| Numbers `focused_bridge_*` (3), `movie_playback` `focused_*` (3) | they tested the removed payload wrappers. The Numbers read and write of the same codec has its own test in `sheet_drawable_properties`; the codec has 7 in `litchi-iwa-protos` |
| Numbers `selected_payload_verification_*`, `physical_compatibility_*` (2) | compat modules are gone |
| Numbers `shared_catalog_reader_pins_*`, Keynote `shared_catalog_keeps_*`, Pages `shared_catalog_ingress_retains_*` | they pinned the identity of an ingress that is gone |
| Numbers `prepared_source_handoff_matches_global_compatibility_projection` | replaced, see above |
| Pages `catalog_body_table_appearance` tests (3), `text_storage` tests (7) | the modules are gone |
| Pages `semantic_prepared_source_releases_physical_source`, `apply_rejects_semantic_prepared_source_*` | the semantic prepared route is gone |
| Keynote `hidden_bridge_redacts_*`, `semantic_only_package_ingress_*` | the bridge and the only way to build that package state are gone |

The fuzz target `numbers_table_cell_comment_reply` no longer calls the compatibility reads.

## 3. Dependencies, dead code and docs

- Removed: `litchi-iwa-style -> litchi-iwa-archive` and `litchi-numbers -> phf`, both unused
  (no `cargo machete` or `udeps` here; I scanned `use` paths and let the compiler confirm). The
  four affected fuzz locks are refreshed.
- Removed from Numbers: `_component_archive`, kept "for the host migration's graph diagnostics",
  and three test fixtures the compatibility suite used.
- Boundary checker: the audits that required a seam are gone. The movie-playback audit lists
  Pages and Keynote only, because it returned `[]` when any owner file was missing and would have
  gone silent when the Numbers module was deleted.
- Glossary: `numbers-full` and `keynote-full` differed from `*-minimal` only by the removed
  feature, and `keynote-minimal` bans one name more. Both are deleted: 70 profiles.
- Feature matrices re-audited against the exit ledger: 19 rows that claimed no support, or
  support the ledger shows is wider, now name the API. The re-audit took the rows marked ❌
  and the ledger's families; it did not recheck each ✅ or 🟡 row's notes. Numbers' row for the
  physical sort seam is removed.
- ADR 0024 and ADR 0028 have a dated amendment. Historical records are unchanged.

## 4. Left, and why

- **`litchi_pages::OpenOptions::__archive_limits`** stays hidden. It served the host, and it also
  serves 23 Pages test files and 11 fuzz targets, which tune archive dimensions
  (`with_archive_limits`) that `SourceLimits` does not name. Removing it needs new public surface,
  which has no non-test consumer, or dropping those budget assertions. The owner can choose.
- **Hidden seams in `litchi-iwa-archive` and `litchi-iwa-detect` that leaves still call**, such as
  `PreparedSource::__into_source_catalog`, `__from_*_with_numbers_semantics`,
  `__classify_application_payload` and `unstable::__prepare_reassembly_with_changes_admitted`.
  These are live coordination between focused crates, not host seams.
- **Archive code that is now unreachable from any other crate**: `LogicalEntryLimits::PAGES_METADATA`,
  `__pages_metadata_sidecars`, `__from_*_with_logical_entry_limits`, and
  `FrozenDirectoryBundle::{open_with_pages_metadata, open_with_pages_metadata_when, open_with_semantic_metadata, into_pages_semantic_parts}`.
  Removing them means editing the live ZIP ingress path, so I left it for its own package.
- **Older dead code behind `allow(dead_code)` in Numbers**: 30 items when the allows are removed
  (`extractor`, `table`, `sheet`, `comments_metadata`). Their reasons do not name the host, so I
  left them.
- **`compatibility_tables_from_bytes`** in Numbers is public and the facade re-exports it.
- **`KEYNOTE_PHYSICAL_SORT_OWNER_ACTIVE`** is a public constant a checker audit uses for activation.
- **Comments that cite "the host's" order or function** in the shared crates are provenance and
  stay. `caption.rs` is iwa-w8's. Unused dependencies my scan listed in other crates (`litchi-odt`
  and others) are outside this package and not confirmed by a compile.

## 5. Byte output

`mac_queue_file_routes` and `mac_queue_native_routes` ran before (at `3001da641b`) and after. Of 28
files, 26 hash the same; the 2 that differ are the two `Identity::generate()` Pages documents, which
also differ between two runs of the same binary. Before: `.coord/iwa-w9/before.sha`, after:
`.coord/iwa-w9/after.sha`.

## 6. Verification

Code head `8d37f51f18`, plus the docs and the one policy-line commit; only matrix text under
`crates/` changed after the tests ran. Logs: `.coord/iwa-w9/gates/`.

| Command | Exit |
|---|---:|
| `cargo fmt --all --check` | 0 |
| `cargo clippy --workspace --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features f] --all-targets -- -D warnings`, no feature and each of the 21 formats | 0 (all 22) |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features`, eight changed crates plus `litchi-ops`, `litchi` | 0 |
| `tools/gate/chunked_cargo_tests.py -p <c> --all-features`, for `litchi-iwa-{archive,detect,table,style,protos}`, `litchi-{numbers,pages,keynote,ops}`, `litchi` | 0 (each); 7,534 passed, 0 failed, 257 ignored |
| `cargo test -p litchi-ops --features <every feature except encryption> --test capabilities --test registry` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `cargo check --locked --bins`, all 20 `crates/*/fuzz` manifests, and `litchi-ops/fuzz --features apply` | 0 (each) |
| `mac_queue_file_routes`, `mac_queue_native_routes` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 (69 packages, 302 declarations, 0 debt items) |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir $CARGO_TARGET_DIR` | 0 (70 profiles) |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_iwork_leaf_public_api.py` | 0 |
| `python3 -m unittest`, the rules' tool modules plus `test_iwa_exit_ledger`, `test_mac_queue`, `test_mac_control_kit`, `test_gate_runner`, `test_check_iwork_leaf_public_api`, `test_check_api_glossary`, `test_api_debt_inventory`, `test_validate_api_disposition`, `test_unified_bench_matrix` | 0 |
| `officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `unified_bench_matrix.py --check`, `iwa_exit_ledger.py --check`, `mac_queue/generate.py --check` | 0 |
| `python3 tools/check_feature_matrix.py` | 0 |

Goldens were not re-blessed: the capability, registry and corpus tests pass unchanged. No fuzz lock
names `litchi-iwa` or the removed `litchi-iwa-style -> litchi-iwa-archive` edge.

First-pass failures, all fixed: `check_api_glossary.py` (it pinned `OwnedSourceCatalog` as an owner);
`check_crate_boundaries.py` (the policy still listed the removed style edge);
`test_unified_bench_matrix` and `unified_bench_matrix.py --check` (the generated matrix had to pick up
this record); 13 tests of `test_check_crate_boundaries` that pinned the seams, before they were
updated.

Not verified: no fuzz target was run, only checked. No Mac run: the byte comparison shows the queue
packages did not change, so the Mac evidence still applies. The full 53-row gate is the coordinator's.
The series builds only at its last code commit, because the feature crosses manifests.
