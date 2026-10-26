# 0372 — W1-C: Keynote slides, drawables and media out of `litchi-iwa`

- Status: partial landing of wave W1-C on branch `uo2/iwa-w1c` (cut from
  trunk `feat/unified-ops-api` at `f35f1c9f55`; trunk merged at
  `112689d58e`, `8f0f66c80f`, `8dd006ef12` and after the W1-B cell-writer
  fix).
- Date: 2026-10-26.
- Plan: [0363](0363-iwa-monolith-exit-plan.md) wave row W1-C (§0), API shape
  §3.2, parity §4; harness and ledger [0369](0369-iwa-exit-parity-harness-and-ledger.md);
  codecs and the D4 gap [0368](0368-iwa-w0b-codecs-formula-graph-table-vocabulary.md);
  owner decision D1 and the [macOS queue](0363-mac-queue.md).
- `performance_claim`: every migrated route is 28–92% cheaper than its host
  route (callgrind, §5); the one miss found (z-order move, +9.4%) was fixed,
  not carried.

## 1. What landed

| Commit | Crate | What |
|---|---|---|
| `fab10d07ab` | `litchi-keynote` | `slide::drawable::Edit::remove` culls data whose final owner was the removed drawable (`DataInfo` + `Data/` member) and releases the object watermark when the removal takes out the object holding it, as the host does (`Diagnostics::removed_data`); `Preset` gains `LeftArrow`, `RightArrow`, `DoubleArrow`, `RegularPolygon`, `Star` (host path sources; shared `PolygonSides`/`StarPoints`/`InnerRadiusRatio` re-exported from `slide::drawable`) |
| `ea3d87a07a` | `litchi-iwa` (host) | two documented host defects fixed (§3): dangling template tags after drawable removal, missing external reference after `add_slide_text_box` |
| `ebbce329e2` | `litchi-iwa` tests, ledger, queue | drawable and slide-lifecycle parity cases; harness learns the leaf's documented policies (root-preview invalidation, save-token advance, deterministic UUIDs of created objects); Keynote drawable schemas registered |
| `e2de754f3b` | `litchi-keynote` | the 0368 gap: `slide::table::number_format::{Number, DecimalPlaces, FixedDecimalPlaces, NegativeStyle, ThousandsSeparator, MAX_DECIMAL_PLACES, NumberError, At, AtError}` are now re-exports of `litchi_iwa_common::table` (`tests/neutral_table_vocabulary.rs` proves nominal identity); no workspace caller needed a change |
| `8ff06a2d5e` | `litchi-iwa` tests | paired callgrind probe `keynote_w1c_cost` |
| `3d006110be` | `litchi-iwa-archive` | `SourceCatalog::__from_shared_bytes_reusing`: a derived candidate reuses the base's parsed archive for byte-identical IWA members (ZIP and budgets still checked in full); parsed components are `Arc`-shared |
| `ba7155292b` | `litchi-keynote` | every structural commit reopens its candidate through the reuse path |
| `a9660091a6` | `litchi-keynote` | layout reassignment: `slide::creation::Edit::set_layout(slide, layout)` (port of host `set_slide_layout`, including the layout-media swap) |
| `4e85d30500` | `litchi-iwa` tests, ledger, queue | layout-reassignment parity on 48 layout pairs |
| `fc445c1aa3` | `litchi-keynote` | `Builder::language`, `Builder::locale` |
| `a760655e92` | `litchi-iwa` tests | builder language/locale field differential |

The host keeps every route as the oracle; no host behaviour changed except
the two defects of §3. No manifest, crate edge or ops dispatch file changed.

## 2. Ledger

22 Keynote rows flipped to `migrated` (25 with the three W0-C merge rows);
614 Keynote rows stay `host-only`. Each migrated route without host native
evidence has a `0363-mac-queue.md` row (12 W1-C rows).

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

Read cases compute host and leaf projections inside one read and fail on
any disagreement (the harness compares each read across the two outputs).

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
- `litchi_iwa_archive::SourceCatalog::__from_shared_bytes_reusing`
  (doc-hidden, unstable).

The Keynote glossary profiles pass unchanged (no new banned names, no new
family; `set_layout` joins the registered `slide::creation` family).

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

Before `3d006110be`/`ba7155292b` the z-order move was **+9.4%** (leaf
120,550,699 vs host 110,203,188): the structural reopen parsed every
component again (22.5M Ir of the 60M leaf route). Reusing the base's parsed
archives for unchanged members fixed it and cut every structural route by a
further ~40M Ir.

## 6. Gaps and follow-ups

| Gap | Rows | Next |
|---|---|---|
| Builds and effects (`add/set/remove_slide_build`, 48 build value constructors) | 52 | leaf build editing family over `build::semantic` values (no donor) |
| Chart lifecycle (`add/duplicate/remove_slide_chart`, kind, direction, geometry) | 8 | port donor `repair/keynote-chart-deploy-20261024` (11k + 2.7k lines) |
| Shape lifecycle beyond add/remove/geometry (duplicate, flip, lines and endpoints, text set/replace/clear, title/caption) and shape style (fill, stroke, shadow, effects, text layout) | 45 | donor `refactor/keynote-shape-text-repair-20261024-r5`; style codecs are W1-E |
| Image lifecycle beyond add/remove/geometry (duplicate, flip, title/caption, properties, data replacement, original size) | 11 | extend `slide::media::lifecycle` to images (integration-line donor) |
| Text-box duplicate, properties and columns | 9 | with W2-C text boxes |
| Table families (cells, rich text, formats, appearance, comments, highlights, formulas, hidden axes, topology, lifecycle) | 147 | after W1-A's engine pieces; the cell-writer divergence of §3 first |
| Builder `language`/`locale` (implemented, `fc445c1aa3`) | 2 (+8 builder rows) | the host builder's random UUIDs defeat a tiered case; the field differential `keynote_builder_language_locale_differential` proves the written fields. The rows stay `host-only` until the harness can normalise builder identities |
| Chart style families | 160 | W2-C |
| Slide text-box rich text and paragraph families | ~140 | W2-C/W1-D |

## 7. Verification

Final run on the merged branch (trunk `31c87001d6` merged), build env of
the coordinator rules (`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`).

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-keynote -p litchi-iwa-archive -p litchi-iwa -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings`, and with `--features keynote` | 0, 0 |
| `cargo test -p litchi-keynote --all-features` (1,181) / `-p litchi-iwa-archive` (214) / `-p litchi-iwa` (1,836, `parity` target 44) / `-p litchi` (209) | 0 / 0 / 0 / 0 |
| `cargo test -p litchi-ops --all-features`: `--lib` (1,022), `--doc` (217), nine groups of ≤20 `--test` targets (1,485) | 0 each |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-keynote -p litchi-iwa-archive` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (71 profiles) | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0; 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_iwa_exit_ledger` (1,289) | 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |
| `python3 tools/officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `unified_bench_matrix.py --check` (after regenerating `results/0289/bench-matrix.md`, which was stale on trunk and now lists this record) | 0, 0, 0 |
| `python3 tools/check_feature_matrix.py` (31 legs) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (301) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |

No capability or corpus-visible behaviour of `litchi-ops` changed (no ops
dispatch arm, D7), so the capability/registry goldens and the OfficeCLI
corpus were not re-blessed; their tests pass unchanged inside the litchi-ops
groups. Both pre-commit size checks printed nothing for every commit.
