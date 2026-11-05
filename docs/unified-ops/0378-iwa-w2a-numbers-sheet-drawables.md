# 0378 — W2-A: Numbers sheet drawables out of `litchi-iwa`

- Status: first merge-ready slice of wave W2-A on branch `uo2/iwa-w2a` (cut
  from trunk `feat/unified-ops-api` at `2d971a95dd`; trunk merged at
  `05b3a011fb` and `485f396b4f`). The wave continues (§6).
- Date: 2026-10-26.
- Plan: [0363](0363-iwa-monolith-exit-plan.md) wave row W2-A (§0), API shape
  §3.2, parity §4; harness and ledger
  [0369](0369-iwa-exit-parity-harness-and-ledger.md); shared engines W1-D
  ([0374](0374-iwa-w1d-shared-text-engine.md): `paragraph_style`,
  `language`, the TSWP rewrite) and W1-E
  ([0375](0375-iwa-w1e-shared-chart-shape-media-engine.md):
  `chart_style_codec`, `shape_properties_codec`, the common chart values);
  owner decision D1 and the [macOS queue](0363-mac-queue.md).
- Donors: none ported in this slice. The Numbers drawable-comments donor
  (`feat/numbers-drawable-comments-focused-rereview-20261024`, 6k lines,
  pre-0351 API) is the next family (§6). Keynote's W1-C drawable family and
  Pages' W1-B drawable properties/text families are the patterns the Numbers
  families follow.
- `performance_claim`: every migrated route is 10–83% cheaper than its host
  route under callgrind ABBA (§5); the one miss found (z-order move, +30%)
  was fixed before landing.

## 1. What landed

All code is in `litchi-numbers` (owned paths `src/package/sheet_{drawable,
chart}*`, `src/sheet/{drawable,chart}.rs`) plus separated hunks in
`package.rs`, `sheet.rs`, `lib.rs` and the shared tool files.

| Commit | Family | What |
|---|---|---|
| `bf5771ef8f` | `sheet::drawable` (structural), `sheet::drawable::properties` | `DrawableSelector::{Position, Of { kind, position }}` (also `litchi_numbers::DrawableSelector`), `Kind`, `Drawable`, `Geometry`, `GeometryFlags`, `Properties`; `Package::sheet_drawable_list`; `edit_sheet_drawables` → `Edit::{move_to, remove}`; `edit_sheet_drawable_properties` → `Edit::{set_geometry, set_position, flip, restore_original_size, set_properties, set_playback}` over text boxes, shapes, images, movies, audio and charts. Private `sheet_drawable_graph` resolves the sheet, classifies its drawable list and publishes in-place rewrites (each touched component cloned from the parsed source, previews and metadata kept) |
| `284722cdfb` | `sheet::drawable::text` | `Package::{sheet_text_box_text, edit_sheet_text_box_text, apply_sheet_text_box_text}`; `Edit::{set(Property), reset(PropertyKind)}` publish the W1-D paragraph-style plans (Rewrite in place; Derive: new variation in the stylesheet's component, storage repointed in place, UUID and external reference registered; Collapse) |
| `b75647c316` | `sheet::drawable` | `Edit::{add_text_box, add_shape(Preset, ..)}`: wire-level port of the host constructors (drawable, two caption stand-ins, storage; theme text-box preset; deterministic UUIDs); `litchi-numbers` now depends on `litchi-iwa-text` (as Pages and Keynote do; boundary policy updated) and `sheet::drawable::text` re-exports the value modules its properties are built from |
| `6eb6e8dbbc` | `sheet::drawable::plain_text` | `Package::{sheet_drawable_plain_text, edit_sheet_drawable_plain_text, apply_sheet_drawable_plain_text}`; `Edit::{replace, insert, delete, set, clear}` on the shared TSWP rewrite |
| `5f3ee173c5`, `f1bfa2b000` | `sheet::chart::style` | `Package::{sheet_chart_style, edit_sheet_chart_style, apply_sheet_chart_style}`; `Property::{PieStartAngle, DonutInnerRadius, RadarStartAngle, SeriesGap3d, Rotation3d, Depth3d, RoundedCorners, LegendFrame}` over the common chart values and the W1-E codec write rules; `Edit::{set, reset}` |
| `cadb92a822` | `sheet::drawable::language` | `Package::{sheet_drawable_text_languages, sheet_drawable_text_language, edit_sheet_drawable_text_languages, apply_sheet_drawable_text_languages}`; `Edit::{set, remove_boundary, reset}` on the W1-D language engine |

Each family is registered in the `numbers-minimal` and `numbers-full`
glossary profiles with its `apply_*` result contract and listed in
`docs/API_GLOSSARY.md` under this record. No native identifier, UUID or
protobuf value appears in a public signature. No litchi-ops dispatch file
changed (D7). The host keeps every route as the oracle; no host behaviour
changed.

## 2. Ledger

134 of the 451 W2-A host rows flip to `migrated` (225 Numbers rows in all):

| Rows | Host routes | Leaf destination | Parity tests | Tier |
|---:|---|---|---|---|
| 30 | text-box/shape/image/movie/chart geometry, shape/image/movie flip, image/movie Original Size, text-box/shape/image/movie/audio properties, audio position, movie/audio playback, and their reads | `sheet::drawable::properties::Edit::*`, `Package::sheet_drawable_properties` | `numbers_sheet_*_{geometry,flip,original_size,properties,position,playback}_parity` | A |
| 3 | `sheet_drawable_order`, `set_sheet_drawable_order`, `move_sheet_drawable` | `Package::sheet_drawable_list`, `Edit::move_to` | `numbers_sheet_drawable_{list_read,order,to_front,backward}_parity` | A |
| 2 | `remove_sheet_{text_box,shape}` | `Edit::remove` | `numbers_sheet_{text_box,shape}_remove_parity` | B |
| 69 | 23 text-box character/paragraph properties × set/reset/read | `sheet::drawable::text::Edit::{set,reset}`, `Package::sheet_text_box_text` | `numbers_sheet_text_box_property_{set,rewrite,reset,read}_parity` | B (set, reset), A (read; in-place rewrites are byte-identical) |
| 3 | `add_sheet_{text_box,rectangle,shape}` | `Edit::{add_text_box,add_shape}` | `numbers_sheet_{text_box,rectangle,shape}_add_parity` (8 presets) | B |
| 6 | `{set,replace,clear}_sheet_{text_box,shape}_text` | `sheet::drawable::plain_text::Edit::*` | `numbers_sheet_{text_box,shape}_text_{replace,set,clear}_parity` | A |
| 16 | chart pie start angle, donut radius, radar start angle, 3D gap/rotation/depth, rounded corners, legend frame (set + read) | `sheet::chart::style::Edit::set`, `Package::sheet_chart_style` | `numbers_sheet_chart_style_{set,reset,read}_parity` | A |
| 5 | text-box language set/remove/reset/reads | `sheet::drawable::language::*` | `numbers_sheet_text_box_language_*_parity` | A |

Only the audio playback rows cite host native evidence
(`audio-playback-native-resaved.numbers`); every other route has a W2-A row
in the macOS queue (25 rows).

## 3. Parity findings

- Inputs: native fixtures where they exist (`image-adjustments-native`,
  `audio-playback-native`, `chart-data-native`, `basic`); text boxes,
  rectangles, movies and charts come from the host's own constructors on
  `basic.numbers` / `chart-data-native.numbers`.
- Tier B differences are the documented leaf policies only, accepted in their
  exact shape by the harness: removals and creations delete the three root
  previews and advance the save token (`leaf_invalidates_root_previews`,
  `leaf_structural_metadata`); created objects get deterministic UUIDs. Object
  identifiers equal the host's in every case.
- In-place edits (frames, properties, playback, z-order, text, chart styles,
  languages, paragraph-style rewrites) are byte-identical to the host,
  including the chart frame route, whose host version re-encodes the chart
  drawable with Prost.
- Host quirk (not fixed, not migrated around): the host's
  `set_sheet_chart_geometry` refuses a chart it has just added to
  `chart-data-native.numbers` ("component UUID map does not cover chart");
  the geometry case therefore uses a chart added to `basic.numbers`.
- The storage repoint of a derived paragraph-style variation keeps the header
  reference in place (the core's transition primitive only appends), so the
  storage header is re-encoded canonically as the host does.

## 4. API notes

- `DrawableSelector` follows the Numbers/Keynote constructor convention
  (`position(Position)`, `index(usize)`, `as_position`, `as_index`) and adds
  kind-scoped constructors; `From<ImageSelector>` and `From<ChartSelector>`
  map the existing selectors onto it.
- Property families take `&mut self` verbs and one `commit` (as Pages'
  drawable families); structural edits stage several operations whose
  selectors resolve against the base snapshot (as Keynote's).
- `cell::text` (W1-A) and `sheet::drawable::text` share the
  `paragraph_style::{Property, PropertyKind}` vocabulary; the value types
  are now reachable through `sheet::drawable::text::{appearance, character,
  font, paragraph}`.

## 5. Performance

Callgrind, release with debug info, CPU 3 pinned (`taskset -c 3`), ABBA
(host, leaf, leaf, host) in one run, inclusive Ir of `w2a_host_route` and
`w2a_leaf_route` (two calls each: open + one operation + commit + serialize;
selectors resolved outside). Probe: `crates/litchi-iwa/tests/parity/numbers_w2a_cost.rs`
(command in its header). Outputs are kept outside the repository in
`(deleted-coord-artifact)`.

| Route (fixture) | Host Ir | Leaf Ir | Δ |
|---|---:|---:|---:|
| image geometry (`image-adjustments-native`) | 121,622,411 | 53,384,441 | −56.1% |
| audio playback (`audio-playback-native`) | 86,152,556 | 45,371,050 | −47.3% |
| text-box properties (authored) | 87,553,279 | 46,125,637 | −47.3% |
| chart geometry (`basic` + chart) | 198,635,780 | 34,511,997 | −82.6% |
| z-order move (authored) | 51,478,625 | 46,133,996 | −10.4% |
| text-box remove | 125,167,116 | 111,078,213 | −11.3% |
| shape remove | 175,160,779 | 111,051,747 | −36.6% |
| text-box alignment set | 264,744,341 | 158,594,866 | −40.1% |
| text-box add | 145,631,476 | 93,580,724 | −35.7% |
| shape add (ellipse) | 202,831,114 | 93,604,527 | −53.9% |
| text-box text set | 77,997,563 | 46,273,996 | −40.7% |
| chart rounded corners (`basic` + chart) | 181,222,274 | 43,298,963 | −76.1% |

The first z-order move went through the structural working set and cost
66.9M Ir against the host's 51.5M (+30%): the working set decodes the
package metadata registry twice (watermark and transition) even when nothing
is added or removed, and re-parses the touched component. In-place edits now
clone the parsed component and skip the registry entirely, which brought the
move to −10.4% and every property route down by ~20M Ir.

## 6. Gaps and follow-ups

| Gap | Host rows | Next |
|---|---:|---|
| Chart style beyond chart level: axes, legend style, series fill/stroke/symbols/labels/trendlines/error bars, per-wedge explosion and label distance (the engine rewrites only existing series non-style objects; the host allocates them on first set), fonts, titles, captions, 3D lighting/bar shape | ~150 | W1-E's series-style write encoders (0375a) are on trunk; per-wedge needs the allocation the host does |
| Chart lifecycle (add, duplicate, remove, kind, direction, data, reads) | 8 | port Keynote's W1-C `chart::lifecycle` approach |
| Text-box lists, drop caps, hyperlinks, highlights, comments and replies, date-time fields, named paragraph styles | ~55 | W1-D engines exist (`list`, `drop_cap`, `smart_field`, `ranged`, `annotation`, `paragraph_style::named`); the publishing follows `sheet::drawable::text` |
| Text-box layout and columns; text-box duplicate | 6 | shape-property codecs |
| Shape style (fill, stroke, shadow, effects, preset, line endpoints/segment, image fill, text layout) and shape duplicate, lines, title/caption | ~32 | W1-E `shape_style_codec` and drawing encoders |
| Image/movie/audio add, remove (data culling), duplicate, data and poster replacement, title/caption | ~35 | needs data-owner removal in the structure metadata transition (W1-A's `structure` refuses data owners today) |
| Drawable comments (`sheet_drawable_comment*`, `sheet_drawables`) | 8 | donor `feat/numbers-drawable-comments-focused-rereview-20261024` |
| Media assets (`media_assets`, `sheet_media_assets`, `extract_media`, `replace_media`) | 4 | W1-E `media::embedded` feed |
| Leaf-only unit coverage | — | `crates/litchi-numbers/tests/sheet_drawables.rs` covers the families on native and leaf-authored drawables; text-box properties are covered there through a leaf-created text box |

## 7. Verification

Final run on the branch with trunk merged (`485f396b4f`), build env of the
coordinator rules (`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`). No
litchi-ops capability or corpus-visible behaviour changed (no dispatch arm,
D7), so the capability/registry goldens and the OfficeCLI corpus were not
re-blessed; their tests pass unchanged inside the litchi-ops groups.

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-numbers -p litchi-iwa -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings`, and with `--features numbers` | 0, 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-numbers` | 0 |
| `cargo test -p litchi-numbers --all-features`, `-p litchi-iwa --all-features` (incl. the `parity` target), `-p litchi --all-features` (3,411 tests) | 0 |
| `cargo test -p litchi-ops --all-features`: `--lib`, `--doc`, nine groups of ≤20 `--test` targets (2,724 tests) | 0 each |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (71 profiles) | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0; 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_iwa_exit_ledger` (1,289) | 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |
| `python3 tools/officecli_capability_delta.py --check`, `officecli_corpus.py --check` | 0, 0 |
| `python3 tools/unified_bench_matrix.py --check` (after regenerating `results/0289/bench-matrix.md` to list this record) | 0 |
| `python3 tools/check_feature_matrix.py` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |

Both pre-commit size checks printed nothing for every commit.
