# 0416 — W4-D: Numbers shape lifecycle, lines and shadows (slice 1)

- Status: slice 1 merge-ready on branch `uo2/iwa-w4d` (cut from trunk
  `12682616f7`; trunk unchanged at the slice head). The package continues
  (§6).
- Date: 2026-11-02.
- Scope (coordinator, after record 0415 §3):
  1. Pages `Builder::body_table` from a from-scratch table graph in
     `litchi_numbers_wire::table_create`;
  2. the host-only Numbers drawable families `drawable.shape` (22),
     `drawable.image` (6), `drawable.movie` (6), `drawable.audio` (1);
  3. `chart.lifecycle` on Numbers (7), Pages (8) and shared (1).
- `performance_claim`: every migrated route is 50–68% cheaper than its host
  route under callgrind ABBA (§4).

## 1. What landed

| Commit | What |
|---|---|
| `5ae1838c64` *(shared)* | `litchi_iwa_style::line_end::{endpoints, set_endpoints, EndpointsError}`: the host's `shape_line_endpoints` and `set_shape_line_endpoints` over a style workspace (own legacy decorations, else the style chain's; swapped for a horizontally flipped path; a default write clears an exclusive variation whose parent is undecorated, else both slots go through the plan layer). Pages `body_drawable_style` drops its copy (−244 lines) |
| `702c8f1ac1` *(shared)* | `litchi_iwa_style::shape_path::{line_segment, segments_match}`: the segment a line's frame and path width draw (and whether its path runs backwards), and the `f32` round-trip comparison. Pages `body_drawable_properties` drops its copies |
| `047f6cdfa0` | Numbers `sheet::drawable`: `Edit::add_line(Segment)` (the existing author with a `Path::Line` source from `shape_path::line_path_source`), `Edit::duplicate` for shapes and lines; `properties::{Snapshot::{preset, line_segment}, Edit::{set_preset, set_line_segment}}` on `shape_path` and the shared line codec; `style::Property::{Shadow, LineEndpoints}` on the engine's `Shadows` and `line_end`. A line-endpoint reset stores the undecorated default (the host's `reset_sheet_line_endpoints`); endpoints of a shape that is not a line are `Refusal::UnsupportedProperty` |
| `9dde6b5e78` | parity `numbers_shape_lifecycle_cases` (15 cases and a cost probe), 19 ledger rows, macOS-queue rows |

No code was copied from the Pages or Keynote leaves: the two pieces the
Numbers families needed that lived in a leaf (line endpoints and line
segments) were lifted into `litchi-iwa-style` first, in their own commits,
and Pages moved onto them. Keynote keeps its own line-endpoint code on its
own workspace type (follow-up, §6).

## 2. Ledger

19 Numbers rows flip to `migrated`; `python3 tools/iwa_exit_ledger.py
--check` passes.

| Rows | Host routes | Destination | Tier |
|---:|---|---|---|
| 3 | `add_sheet_line`, `add_sheet_line_with_{endpoints,style}` | `sheet::drawable::Edit::add_line` (the combined routes compose `edit_sheet_drawable_style`, as the host composes its setters) | B |
| 1 | `add_sheet_shape_with_fill` | `Edit::add_shape` + `style` fill | B |
| 1 | `duplicate_sheet_shape` | `Edit::duplicate` | B |
| 3 | `{,set_,reset_}sheet_line_endpoints` | `Package::{sheet_drawable_style, edit_sheet_drawable_style}` | read A, writes B |
| 3 | `{,set_,reset_}sheet_shape_shadow` | the same | read A, writes B |
| 2 | `{,set_}sheet_shape_preset` | `properties::{Snapshot::preset, Edit::set_preset}` | A |
| 2 | `{,set_}sheet_line_segment` | `properties::{Snapshot::line_segment, Edit::set_line_segment}` | A |
| 4 | `sheet_{shapes,images,movies,audio}` | `Package::sheet_drawable_properties` (with `sheet_drawable_list`) | A |

Every route has a W4-D row in `0363-mac-queue.md` (no host native
evidence). Numbers host-only rows: 114 → 95.

## 3. Parity findings

- **Identifier order.** Line creations and shape copies allocate the
  host's identifiers in the host's order; only the documented UUID and
  save-token policies make them Tier B.
- **Presets and segments** are byte-identical to the host (Tier A).
- **Theme shadows.** As on Keynote (record 0372), most corpus Numbers
  themes keep enabled-shadow fields on their disabled shadow; the host
  refuses every shadow read and write on them, and the leaf refuses the
  same way (`numbers_sheet_shape_shadow_inconsistent_theme_refusal_parity`).
  Shadow parity runs on `chart-grid-authored-native.numbers`, whose theme
  shadow both sides decode.
- **Line copies.** The host's `duplicate_sheet_shape` compared the moved
  copy's segment with the source's and refused every line (the defect fixed
  for Pages in `428c32a13c`). Slice 2 fixes the host check the same way
  (the expected segment is the source's moved by the duplicate offset;
  macOS-queued as a host fix), and the duplicate parity case now copies
  both the rectangle and a line.

## 4. Performance

Callgrind on a release build with debug info, CPU 3, ABBA (two calls of
each route: open, edit, commit, serialize); probe
`numbers_shape_lifecycle_cases::numbers_w4d_shape_route_cost_abba`.

| Route | Host Ir | Leaf Ir | Δ |
|---|---:|---:|---:|
| line add | 252,637,409 | 94,127,282 | −62.7% |
| shape duplicate | 261,032,898 | 94,668,347 | −63.7% |
| line endpoints set | 289,874,294 | 145,262,611 | −49.9% |
| shadow set | 67,633,903 | 30,581,683 | −54.8% |
| preset set | 132,936,876 | 46,642,939 | −64.9% |
| line segment set | 146,562,352 | 46,565,713 | −68.2% |

```sh
CARGO_PROFILE_RELEASE_DEBUG=1 cargo test --release -p litchi-iwa --test parity --all-features --no-run
W4D_ROUTE=<route> taskset -c 3 valgrind --tool=callgrind target/release/deps/parity-<hash> \
  --ignored --exact numbers_shape_lifecycle_cases::numbers_w4d_shape_route_cost_abba --test-threads=1
callgrind_annotate --inclusive=yes callgrind.out.<pid> | grep w4d_
```

The raw outputs stay in the session scratch (`.coord/iwa-w4d/cg/`).

## 5. Verification

Run on `9dde6b5e78` (trunk `12682616f7`, unchanged), build env of the
coordinator rules (`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`), test
executables pruned after each chunk. No litchi-ops dispatch arm changed, so
the capability/registry goldens and the OfficeCLI corpus were not
re-blessed. Every row exited 0:

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-iwa-style -p litchi-pages -p litchi-numbers -p litchi-iwa -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings`; the same with `--features numbers` and `--features pages` | 0, 0, 0 |
| `cargo test --all-features` of litchi-iwa-style, litchi-pages, litchi-numbers, litchi-iwa (parity included) and litchi | 0 each |
| `cargo test -p litchi-ops --all-features`: `--lib`, `--doc`, nine groups of ≤20 `--test` targets | 0 each |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-iwa-style -p litchi-pages -p litchi-numbers -p litchi-iwa` | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0, 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_iwa_exit_ledger` | 0 |
| `python3 tools/iwa_exit_ledger.py --check`; `officecli_capability_delta.py --check`; `officecli_corpus.py --check`; `check_feature_matrix.py` | 0 each |
| `python3 tools/unified_bench_matrix.py --check` | 0 after regenerating the matrix for this record (the first run reported it stale because of this record) |

Both pre-commit size checks printed nothing for every commit.

## 6. Left for the continuation

| Item | Rows | Next |
|---|---:|---|
| Numbers shape image fill | 1 | `style::Edit::set_image_fill` through the Numbers structure's data-owner transition (the Pages path of record 0415, on `structure::metadata::Transition`) |
| Numbers shape, image and movie titles and captions | 15 | the caption graph seam (Pages `body_drawable_caption`, Keynote `slide_movie_caption`): lift the graph authoring and stand-in swap into a shared engine, then a Numbers `sheet::drawable::caption` family |
| `chart.lifecycle` | 16 | lift Keynote's chart graph author (`slide_chart_lifecycle/create.rs`, the host's `charts::source`) into a shared engine, then Numbers and Pages add/duplicate/remove/data/direction/kind/list |
| Pages `Builder::body_table` | 1 | from-scratch table graph in `litchi_numbers_wire::table_create` |
| Keynote line endpoints | — | move `slide_drawable/shape_style.rs` onto `line_end::{endpoints, set_endpoints}` once its workspace exposes the shared `graph::Workspace` |
