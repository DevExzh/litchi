# 0375 — W1-E: the shared chart, shape and media engine

- Status: landed on branch `uo2/iwa-w1e` (cut from trunk
  `feat/unified-ops-api` at `58fda6c0c8`; trunk merged at `78ed1cf99b`,
  `eb2961e75a` and `39a2c06629`).
- Date: 2026-10-26.
- Plan: [0363](0363-iwa-monolith-exit-plan.md) wave W1-E (§0, §3.1, §3.2);
  substrate [0368](0368-iwa-w0b-codecs-formula-graph-table-vocabulary.md),
  harness and ledger [0369](0369-iwa-exit-parity-harness-and-ledger.md);
  owner decision D1 and the [macOS queue](0363-mac-queue.md).
- Donors: none ported. The `repair/keynote-chart-deploy-20261024` "shared
  parts" are the Keynote chart lifecycle, which W1-C landed in
  `litchi-keynote` (`chart::lifecycle`); everything here is ported from the
  host directly.
- `performance_claim`: every migrated route is cheaper than its host route
  (callgrind ABBA, §5); none misses the bar.

## 1. What landed

| Commit | Crate | What |
|---|---|---|
| `08b01d0e62` | `litchi-iwa-protos` | `chart_style_codec`: strict bounded reader and batched rewriter of the generated property extension (field 10000) of the eight TSCH style archives, and the chart style-reference projection |
| `6e7df132ee` | `litchi-iwa-common` | chart geometry values: `chart::pie::{StartAngle, WedgeExplosion, WedgeIndex, LabelDistance}`, `chart::donut::InnerRadius`, `chart::radar::StartAngle`, `chart::corner::{Radius, Rounding}`, `chart3d::{Depth, Scale, Rotation, SeriesGap}` |
| `9eb3a5de8e` | `litchi-iwa` | host chart values wrap the common values; 48 Tier A chart-property cases; 11 ledger rows |
| `de1fae3a36` | `litchi-iwa-protos` | `chart_style_codec::property`: the library write rule of each chart property (coordinator request) |
| `a04179e22e`, `4efd9dbccf` | common, `litchi-iwa` | `media::{Asset, Catalog, Stats}`, `media::asset::Limits`; host media values delegate; 3 read cases; 11 rows migrated, 5 retired |
| `80678653ed`, `5668dc9879`, `bd53076038` | protos, common, `litchi-iwa` | `shape_style_codec::{line, rewrite_line}`, `shape::line::Segment`; host `LineSegment` delegates; 3 Tier A line-move cases; 2 rows |
| `47b580bc3f` | `litchi-iwa` (tests) | line and media cost probes; media engine reads the member table only |
| `7ce546e1a3` | `litchi-iwa-protos` | `shape_style_codec::native::{color, stroke, fill, shape_style}`, `shape_style_codec::style_reference`, `chart_style_codec::{parent, Properties::{message, varint, bytes}, StyleReferences::series_style}` |
| `7529b2a6ed` | `litchi-iwa-common` | `chart::series` (strokes, fills, connection lines, symbols, symbol fills, outlines, trendlines and the per-kind field families), `shape::fill::{Image, ImageData, ImageTechnique, Fill}`, `shape::native` (semantic reading of stored drawing values) |
| `2bfb89501a` | `litchi-pages`, `litchi-numbers`, `litchi-keynote` | identical `drawing_native.rs` field pass (decision (a), §4) |
| `9dd42f6424`, `ddc9264cd2` | `litchi-iwa` | host series and image-fill values delegate; 24 Tier A read cases; 26 rows; read cost probe |
| `613d244fb2`, `6eef1e9170`, `9ce718d818` | protos, common, `litchi-iwa` | `chart_style_codec::{legend_frame, rewrite_legend_frame}`, `chart::legend::{Coordinate, Extent, Rect, Frame}`; 6 Tier A cases; 3 rows |
| `303039d1c9`, `44cb59b0de`, `9b326c8e45` | common, leaves, `litchi-iwa` | `media::embedded::{AssetId, Embedded, Inventory}`; identical `media_native.rs` feed per leaf; embedded-media and extraction read cases; 5 rows migrated, 20 retired |
| `cf1a11016e` | `litchi-iwa` (tests) | the engine reads each IWA component once; embedded-media probe |
| `be4f1e0c69` | `litchi-numbers` | two redundant intra-doc link targets on trunk that failed `rustdoc -Dwarnings` and the glossary gate (doc-only) |
| this commit | docs | this record |

No `Cargo.toml` and no crate edge changed (`check_crate_boundaries.py` 0).
The leaf crates gain only separate W1-E `lib.rs` hunks and two new files
each. The protos build-script inventory registers the two raw codecs.

## 2. The engine

### 2.1 Chart styles (`litchi_iwa_protos::chart_style_codec`)

- `decode(source, archive, limits) -> Properties` validates the outer style
  payload (the extension must be unique, length-delimited and canonically
  framed) and borrows the extension; `float`, `bool`, `int32`, `varint`,
  `bytes`, `message` and `vector` read single slots strictly (duplicate and
  wire-type errors).
- `rewrite(source, archive, writes, empty, limits) -> Rewrite` applies a batch
  of `Write`s with the host's byte semantics: a present field is replaced in
  place (key bytes kept, the length prefix kept when the length is
  unchanged), a cleared field removed, a missing field appended, and a
  missing extension created with its fields in field-number order (what the
  host's Prost encode wrote). The conditional writes carry the host's
  presence rules: `FloatUnlessAbsentDefault`, `BoolUnlessAbsentDefault`
  (rounded corners), `VectorComponents` (3D rotation), `VectorDepth` (3D
  depth: `z = ratio × x` or `y` of the stored vector). `EmptyExtension::
  {Keep, Remove}` covers the per-wedge properties, whose emptied extension
  the host removes.
- `property::{pie_start_angle, donut_inner_radius, radar_start_angle,
  series_gap_3d, rotation_3d, depth_3d, rounded_corners,
  pie_wedge_explosion, pie_label_distance}` take the stored natives of the
  common values and return a `PropertyWrite` (archive, writes, policy). The
  W2 chart families call these; the parity cases do too.
- `references(chart_drawable, limits)` projects the chart's style, per-axis,
  per-series (sparse, with `series_style(index)` = private override else
  theme style) and paragraph-style references without decoding the grid;
  `parent(style)` reads `TSS.StyleArchive.parent`;
  `legend_frame`/`rewrite_legend_frame` read and set
  `TSCH.ChartArchive.legend_frame`.
- Errors are typed (`Malformed`, `DuplicateField`, `WireType`,
  `InvalidScale`, `InvalidField`, `Limit { kind, observed, maximum }`,
  `InvalidLimit`, `Allocation`); `Limits` builders are checked. No Prost
  or Buffa ownership in production code (registered raw codec, production
  ratchet); Prost-differential unit tests.

### 2.2 Shapes (`litchi_iwa_protos::shape_style_codec`)

`line`/`rewrite_line` read and move a native straight line (geometry
components, Bézier natural width, `lineTo` x patched in place);
`style_reference` reads a shape's style; `native::{color, stroke, fill,
shape_style}` decode `TSP.Color`, `TSD.StrokeArchive`, `TSD.FillArchive`
(color, gradient, image) and the parent/fill/stroke slots of
`TSWP.ShapeStyleArchive` into plain stored fields.

### 2.3 Values (`litchi_iwa_common`)

`chart::{pie, donut, radar, corner, legend, series}`, `chart3d`,
`shape::{fill, line::Segment, native}`, `media::{asset, embedded}`. The
semantic rules of the native representation live here: ranges and
defaults, the 3D depth-ratio math and default scene vectors, the line
geometry math, the per-kind series field families and their read rules
(connection-line visibility, symbol-fill switches, trendline flags), the
stored-drawing conversions (`shape::native`), media discovery limits and the
embedded-media inventory. Errors are typed enums; host prefixed names
(`Chart*`, `ShapeImage*`, `ShapeFill`, `LineSegment`, `Media*`,
`EmbeddedMediaAsset`) are banned in the `iwa-common` glossary profile.

### 2.4 Host delegation

The host value types now wrap the common values (constructors, constants,
accessors), with their API, error texts and `Debug` output unchanged. The
host wire patches, chart graph resolution and media package scan stay as
the parity oracle; they are deleted with the host in W3.

## 3. Parity

All cases are in `crates/litchi-iwa/tests/parity/` and run on native-saved
fixtures (`chart-data-native.{key,numbers,pages}`,
`slide-movie-geometry-source-native.key`, `image-adjustments-native.pages`,
`table-relocation-native.numbers`), with a chart, line or shape the host
added where the fixture has none. The engine route is the engine the
families wrap: codec projection and rewrite, common values, one-component
exact-source reassembly through `litchi_iwa_archive`.

| File | Cases | Tier | Routes |
|---|---:|---|---|
| `chart_style_cases.rs` | 48 | A (byte-identical ZIP) | pie start angle, donut inner radius, radar start angle, 3D series gap, rotation, depth (bar, column, pie), rounded corners — set and reset each; per-wedge explosion and label distance |
| `chart_style_cases.rs` | 6 | A | legend frame set and reset |
| `shape_cases.rs` | 3 | A | straight-line move |
| `series_style_cases.rs` | 24 | A, read-only | per-series strokes, fills, connection lines, symbols, symbol fills, symbol outlines, trendlines; shape image fill |
| `media_cases.rs` | 9 | A, read-only | `Data/*` catalog and statistics; embedded-media inventory; extraction |

Every case runs on Keynote, Numbers and Pages. Ledger: 58 rows migrated and
25 retired in the W1-E paths (`charts/**`, `shapes/**`, `media/**`); 57
W1-E rows in the macOS queue (no host route in scope has a native receipt).
The series and image-fill cases are read routes: the host styles the chart
or shape, both routes reopen it unchanged and the reads compare; the write
side of those properties is §6.

## 4. The leaf field passes (decision (a))

The crate graph keeps `litchi-iwa-common` and `litchi-iwa-protos` apart
(neither has a dependency), so two plain field passes sit in each leaf
crate, identical in `litchi-pages`, `litchi-numbers` and `litchi-keynote`:

- `drawing_native.rs` — `shape_style_codec::native` → `shape::native`
  (`color`, `stroke`, `gradient`, `image`, `fill`);
- `media_native.rs` — `package_metadata_media_codec` and the IWA objects'
  data references → `media::embedded::Inventory`.

They are crate-private in the supported build and `#[doc(hidden)] pub` under
`internal-iwork-source` for the parity harness. Each copy's unit tests
compare the three files (`include_str!`) and round-trip one value.
**Triplication noted for W3**: fold them when a shared crate that sees both
crates exists.

## 5. Performance

Callgrind, release with debug info, CPU 3 pinned, instruction counts (Ir),
each probe running host and engine routes in ABBA order over the three
formats (ignored tests; toggle `*<module>::{host,leaf}_<route>*`).

| Probe (route) | Host | Engine | Δ |
|---|---:|---:|---:|
| `chart_style_set_cost_abba` (pie start angle, column depth, rounded corners; 18 runs each) | 7,274,120,706 | 355,679,917 | −95.1% |
| `line_move_cost_abba` | 596,992,655 | 117,668,927 | −80.3% |
| `media_catalog_cost_abba` (open, list, queries) | 31,975,746 | 30,879,607 | −3.4% |
| `embedded_media_cost_abba` | 169,850,518 | 127,474,215 | −24.9% |
| `series_style_read_cost_abba` (strokes, fills, trendlines) | 3,308,933,890 | 329,007,740 | −90.1% |

The host write routes include their own reopen-and-verify; the engine routes
include neither the families' selector resolution nor their candidate
admission, which W2 adds and measures. Two engine regressions found and
fixed before landing: feeding the media catalog from `SourceCatalog` was
+645% (it parses every component; the catalog needs the member table), and
the embedded inventory was +36% while the engine parsed every component
twice.

```sh
CARGO_PROFILE_RELEASE_DEBUG=1 cargo test --release -p litchi-iwa --test parity --no-run
taskset -c 3 valgrind --tool=callgrind '--toggle-collect=*chart_style_cases::host_style_route*' \
  target/release/deps/parity-<hash> --ignored --exact \
  chart_style_cases::chart_style_set_cost_abba --test-threads=1
# likewise leaf_style_route; shape_cases::{host,leaf}_line_route with line_move_cost_abba;
# media_cases::{host,leaf}_{media,embedded}_route; series_style_cases::{host,leaf}_series_route
```

## 6. Handed to other waves

| Capability (host) | Ledger rows still `host-only` | Engine ready | Owner |
|---|---|---|---|
| chart font (`ChartFont::named`, `ChartFontSize::from_points`) | `charts/font.rs` ×2 | values already in `litchi_iwa_text::font::Font` and the text point size; chart paragraph styles via `StyleReferences::paragraph_styles`; char-property rewrite in `litchi_iwa_text_wire` | W2-A/B/C chart style families |
| chart drawable codec, reference lines (`IWorkChartArchive::{decode, encode, new, reference_lines}`) | `charts/archive.rs` ×4 | style references, legend frame; reference-line values in `common::chart::reference_line` | W1-C (`chart::lifecycle` landed), W2 chart families |
| media writes (`media::Draft::{replace, insert_unreferenced, remove_unreferenced, set_max_replacement_bytes}`) | `media/package.rs` ×4 | `media::embedded::Inventory`, `media::Catalog`, `package_metadata_media_codec` rewrites | W1-B, W1-C (`edit_slide_media_data`), W2-A media families |
| writes of series strokes, fills, symbols, symbol fills/outlines, connection lines, trendlines; shape fill/stroke/shadow/effects/line-end/text-layout writes | none (values migrated with read parity) | readers, families and `chart_style_codec::rewrite`; the encoders and the style-variation graph engine continue in [0375a](0375a-iwa-w1e-style-variation-write-engine.md) | W1-E (0375a), then the W2 chart and drawable families |
| `drawable_order.rs`, `table_hidden_axes.rs`, `image_caption.rs`, chart graph clone | none (private modules) | not ported | W2-A/B/C (drawable order and captions), W1-A/B/C (hidden axes, table owners), W1-C (chart clone, landed in `chart::lifecycle`) |

## 7. Verification

Build env of the coordinator rules (`CARGO_BUILD_JOBS=6`,
`RUST_TEST_THREADS=8`, OOM-capped cargo shim); trunk `39a2c06629` merged.

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-iwa-protos -p litchi-iwa-common -p litchi-iwa -p litchi-pages -p litchi-numbers -p litchi-keynote -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-pages -p litchi-numbers -p litchi-keynote --all-targets -- -D warnings` (no features) | 0 |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings`, and with `--features numbers`, `pages`, `keynote` | 0, 0, 0, 0 |
| `cargo test --all-features -p litchi-iwa-protos -p litchi-iwa-common -p litchi-pages -p litchi-numbers -p litchi-keynote` (4,360 tests) | 0 |
| `cargo test --all-features -p litchi-iwa` (1,996 tests, `parity` target included) | 0 |
| `cargo test -p litchi-ops --all-features`: `--lib` (1,022), `--doc`, seven groups of 25 `--test` targets (1,485) | 0 each |
| `cargo test -p litchi --all-features` (209) | 0 |
| `cargo clippy -p litchi-ods --all-targets --all-features -- -D warnings`; `cargo test -p litchi-ods --all-features` (1,783) | 0, 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-iwa-protos -p litchi-iwa-common -p litchi-pages -p litchi-numbers -p litchi-keynote` | 0 (after `be4f1e0c69`) |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (71 profiles) | 0 (after `be4f1e0c69`) |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0, 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_iwa_exit_ledger` (1,289) | 0 |
| `python3 tools/officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `unified_bench_matrix.py --check`, `iwa_exit_ledger.py --check` | 0, 0, 0, 0 |
| `python3 tools/check_feature_matrix.py` (31 legs) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (301) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |

No capability, registry or OfficeCLI-visible behaviour changed, so the
goldens and the corpus were not re-blessed (their tests pass inside the
litchi-ops chunks). Both pre-commit size checks printed nothing for every
commit.
