# 0418 — W4-E: shared chart author, Numbers chart lifecycle, Keynote on the shared modules

- Status: slices 1 and 2 on branch `uo2/iwa-w4e` (cut from trunk
  `b667dc4246`; trunk `2604268add` merged). Slice 2 (§7) adds Pages chart
  creation and frame writes. The rest is left (§6).
- Date: 2026-11-02.
- Scope (coordinator, after record 0416a §5): chart add, duplicate and
  remove on Numbers and Pages over a shared chart author; Pages chart
  frame writes; the shared `document.rs charts` row; Numbers shape image
  fill; Numbers titles and captions; Pages `Builder::body_table` from
  scratch; Keynote onto the shared `line_end` and `chart_lifecycle`.
- `performance_claim`: every migrated route costs 53–85% fewer
  instructions than its host route under callgrind ABBA (§4).

## 1. What landed

| Commit | What |
|---|---|
| `20b387a5be` *(shared)* | `litchi_iwa_style::chart_lifecycle::author::{Ids, Profile, Request, objects, series_count}`: the source-free chart graph (the host's `charts::source`), lifted from Keynote's `slide_chart_lifecycle/create.rs`. Identifiers in the host's order; every payload in tag order; the host's application profile (Numbers: chart mediator, thirty paragraph styles; Pages: owned preset, thirty; Keynote: twenty-five). The owner writes the `TSD.DrawableArchive`, inserts the objects and registers stylesheet, theme, attachment and registry edges |
| `eb966b1927` | Keynote authors new charts through the shared author; its copy (≈400 lines) is deleted |
| `e584786d95` | Keynote chart kind and direction read and write through `chart_lifecycle::{values, with_kind, with_direction}`; Keynote line-end payloads through `line_end::{encode, decode}` (≈250 lines deleted). Endpoint resolution stays on Keynote's own structure workspace, which is not the shared `graph::Workspace` that `line_end::{endpoints, set_endpoints}` take |
| `22b99c0848`, `b95ef0bb20` | Numbers `sheet::drawable::Edit::add_chart(Kind, ChartData, Geometry)`; `Edit::{duplicate, remove}` also take charts; `Refusal::InvalidChart`. The graph lives in the CalculationEngine component as the host builds it; styles join the document stylesheet and the preset the theme's chart presets (created on first use); edges, UUIDs, data owners and the identifier watermark go through the structural `Transition` |
| `bdb4972464` | parity `chart_create_cases` (4 cases, a refusal case, a callgrind probe); 3 ledger rows; macOS-queue rows |

## 2. Ledger

3 Numbers rows flip to `migrated`; `python3 tools/iwa_exit_ledger.py --check`
passes.

| Host route | Destination | Tier |
|---|---|---|
| `add_sheet_chart` | `sheet::drawable::Edit::add_chart` | B |
| `duplicate_sheet_chart` | `sheet::drawable::Edit::duplicate` | B |
| `remove_sheet_chart` | `sheet::drawable::Edit::remove` | B |

Tier B: the host re-encodes the stylesheet and theme with Prost (the leaf
appends the references in place) and draws random UUIDs (the leaf's are
deterministic, the documented policy). Every other changed object equals
the host's byte for byte; the 25 changed objects of the add case agree on
every known field. None of these host routes has native evidence, so each
has a W4-E row in `0363-mac-queue.md`.

## 3. Parity findings

- The leaf allocates the host's identifiers in the host's order for all
  three routes (21 for an add, the graph's size for a copy).
- A chart copy refuses a private reference outside the chart, preset and
  mediator reference paths (`Refusal::UnsupportedDependency`), as the
  host's "unrecognized private reference" check does.
- A removal that leaves the CalculationEngine without charts withdraws its
  stylesheet and paragraph-style edges, as the host does.

## 4. Performance

Callgrind on a release build with debug info, CPU 3, ABBA, two calls of
each route (open, edit, commit, serialize); probe
`chart_create_cases::w4e_chart_create_cost_abba`, selected with
`W4E_CHART_ROUTE`. Raw outputs: `(deleted-coord-artifact)`.

| Route | Host Ir | Leaf Ir | Δ |
|---|---:|---:|---:|
| Numbers chart add | 451,237,881 | 210,709,187 | −53.3% |
| Numbers chart duplicate | 725,297,806 | 215,565,086 | −70.3% |
| Numbers chart remove | 670,047,561 | 238,902,135 | −64.3% |
| Keynote chart add (shared author) | 931,235,036 | 139,929,485 | −85.0% |

```sh
CARGO_PROFILE_RELEASE_DEBUG=1 cargo test --release -p litchi-iwa --test parity --all-features --no-run
W4E_CHART_ROUTE=<route> taskset -c 3 valgrind --tool=callgrind target/release/deps/parity-<hash> \
  --ignored --exact chart_create_cases::w4e_chart_create_cost_abba --test-threads=1
callgrind_annotate --inclusive=yes callgrind.out.<pid> | grep w4e_chart_
```

## 5. Verification

Run on `b95ef0bb20` (trunk `b667dc4246`, unchanged) with this record in the
working tree, build env of the coordinator rules (`CARGO_BUILD_JOBS=6`,
`RUST_TEST_THREADS=8`), one cargo command at a time, test executables pruned
after each chunk (script and log: `(deleted-coord-artifact)`, `gate-s1.log`).
No litchi-ops dispatch arm changed, so the capability/registry goldens and
the OfficeCLI corpus were not re-blessed. Every row exited 0:

- `cargo fmt --all --check`;
- `cargo clippy -p litchi-iwa-style -p litchi-numbers -p litchi-keynote -p litchi-iwa -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings`;
- `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings`, and the same with `--features numbers` and `--features keynote`;
- `cargo test --all-features` of litchi-iwa-style, litchi-numbers, litchi-keynote, litchi-iwa (parity included) and litchi;
- `cargo test -p litchi-ops --all-features`: `--lib`, `--doc` and nine groups of ≤20 `--test` targets;
- `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features` of the four changed crates;
- `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify`;
- `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py`;
- `crates/pyo3-litchi/tests/run_tests.sh`;
- `cd tools/perf-baseline && cargo check --locked --all-targets`;
- `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_iwa_exit_ledger`;
- `python3 tools/iwa_exit_ledger.py --check`, `officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `check_feature_matrix.py`;
- `python3 tools/unified_bench_matrix.py --check`: 0 after regenerating the
  matrix for this record (the first run reported it stale because of this
  record, as for 0416).

Both pre-commit size checks printed nothing for every commit.

## 6. Left

| Item | Rows | Next |
|---|---:|---|
| Pages chart duplicate and remove | 2 | Pages `DrawableSelector` has no chart variant; add chart verbs (`chart::lifecycle::Edit::{duplicate, remove}` or a `DrawableSelector::Chart`) over the body-media graph proof: clone or delete the private graph, the attachment and U+FFFC, unregister styles and preset (as the Numbers module `sheet_drawable_chart` does) |
| shared `document.rs charts` | 1 | aggregate read: retire with the three leaf chart lists, or a `litchi-iwa-structured` aggregate |
| Numbers shape image fill | 1 | `style` image fill through `litchi_iwa_style::shape::{ReleasingFills, stored_fill}` (Pages 0413b/0415) |
| Numbers shape, image and movie titles and captions | 15 | reassigned by the coordinator to W4-F (with chart captions) |
| Pages `Builder::body_table` | 1 | from-scratch table graph in `litchi_numbers_wire::table_create` |
| Keynote line endpoints | — | `line_end::{endpoints, set_endpoints}` once Keynote's shape-style workspace is the shared `graph::Workspace` |

## 7. Slice 2 — Pages chart creation and frame

| Commit | What |
|---|---|
| `99d19c56a3`, `bfa650f286` | Pages `drawable::lifecycle::Edit::add_chart(TextPosition, Kind, ChartData, Point, Size)` and `Refusal::InvalidChart` (submodule `body_media_lifecycle/chart.rs`): the shared author's graph with the Pages profile in the body component, attachment, U+FFFC, z-order created on first use, styles appended to the stylesheet and the preset to the theme's chart presets, UUIDs at the body component and the host's cross-component edges. `chart::lifecycle::Edit::set_geometry(Geometry)` and `Refusal::InvalidGeometry`: the frame parts that differ are rewritten in place and a moved chart's attachment offsets follow (horizontal from the body's left margin, vertical from the page). `chart::lifecycle::Patch::is_noop` is now byte-based, so a frame-only patch and its inverse replay |

| Host route | Destination | Tier |
|---|---|---|
| `add_body_chart` | `drawable::lifecycle::Edit::add_chart` | B |
| `set_body_chart_geometry` | `chart::lifecycle::Edit::set_geometry` | B |

Both are macOS-queued (no host native evidence). Ledger: 2 Pages rows flip.

| Route | Host Ir | Leaf Ir | Δ |
|---|---:|---:|---:|
| Pages chart add | 970,320,609 | 172,192,318 | −82.3% |
| Pages chart frame | 1,930,034,438 | 44,007,224 | −97.7% |

Same probe (`W4E_CHART_ROUTE=pages_add|pages_geometry`).
