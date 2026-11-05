# 0416a — W4-D slice 2: chart kind, direction, data and lists; Numbers line copies

- Status: slice 2 of [0416](0416-iwa-w4d-numbers-shapes-charts-body-table.md)
  on branch `uo2/iwa-w4d`. Trunk merged at `a18d806e4a`, which includes the
  slice 1 merge `831c3a368b` and W4-C slice 5 `10af106c18`.
- Date: 2026-11-02.
- `performance_claim`: every migrated route is 48–97% cheaper than its host
  route under callgrind ABBA (§4).

## 1. What landed

| Commit | What |
|---|---|
| `dd573c13ee` *(host fix)* | `Draft::duplicate_sheet_shape` compared a copied line's segment with the source's unmoved segment, so it refused every line copy. It now expects the source segment moved by the duplicate offset; a `Preserve` placement keeps it unmoved. This is the Pages fix `428c32a13c` applied to Numbers, and it is macOS-queued |
| `8a1284e6a9` | `numbers_sheet_shape_duplicate_parity` now copies both the rectangle and a host-added line, so the leaf line copy is compared with a working host |
| `13477f7056` *(shared, new file)* | `litchi_iwa_style::chart_lifecycle::{values, with_kind, with_direction, assignable_kind}`. They read and rewrite `chart_type` and `series_direction` in place in the chart drawable's `TSCH.ChartArchive` extension. `chart_style` (W4-C) is not touched |
| `a25c3eb652` | Numbers `sheet::chart::lifecycle`: `Package::{sheet_chart_direction, edit_sheet_chart_lifecycle, apply_sheet_chart_lifecycle}` and `Edit::{set_kind, set_direction}`, published in place through the sheet-drawable graph |
| `088284b1c4` | Pages `chart::lifecycle`: `Package::{body_chart_direction, body_chart_geometry, edit_body_chart_lifecycle, apply_body_chart_lifecycle}` and `Edit::{set_kind, set_direction}`. The glossary (`API_GLOSSARY.md`, `api_glossary.json`) registers both families |
| `947108d333` | `chart_lifecycle_cases`: 8 parity cases and a cost probe; 8 ledger rows; macOS-queue rows |

Keynote's `chart::lifecycle` still sets the same two fields on its own
workspace. Moving it onto `chart_lifecycle` is a follow-up (§5).

## 2. Ledger

8 rows flip to `migrated`, and `python3 tools/iwa_exit_ledger.py --check`
passes.

| Format | Host routes | Destination | Tier |
|---|---|---|---|
| Numbers | `set_sheet_chart_{kind,direction}` | `sheet::chart::lifecycle::Edit::{set_kind, set_direction}` | B |
| Numbers | `set_sheet_chart_data` | `sheet::chart::data::Edit::set` (an existing family; no parity case existed before) | B |
| Numbers | `sheet_charts` | `Package::sheet_chart_metadata` with `sheet_chart_direction`, `sheet_chart_data`, `sheet_chart_arrangement` and the drawable frame | A |
| Pages | `set_body_chart_{kind,direction}` | `chart::lifecycle::Edit::{set_kind, set_direction}` | B |
| Pages | `set_body_chart_data` | `chart::data::Edit::set` | B |
| Pages | `body_charts` | `Package::body_chart_metadata` with `body_chart_direction`, `body_chart_geometry`, `body_chart_data`, `body_chart_arrangement` | A |

The writes are Tier B because the host decodes the chart drawable with
Prost and re-encodes it, while the leaf rewrites only the two fields. The
data cases replace the grid with one of the same shape. The host's
reshaping fallback, `set_*_chart_data_full_replace`, is not compared (§5).

## 3. Parity findings

- Every case agrees with the host on known fields and on every leaf and
  host read.
- The host's `body_charts` list carries each chart's frame. Pages had no
  leaf read for a chart frame, so `body_chart_geometry` adds one.

## 4. Performance

Callgrind on a release build with debug info, CPU 3, ABBA, with two calls
of each route (open, edit, commit, serialize). The probe is
`chart_lifecycle_cases::w4d_chart_route_cost_abba`, selected with
`W4D_CHART_ROUTE`; the raw outputs are in `(deleted-coord-artifact)`.

| Route | Host Ir | Leaf Ir | Δ |
|---|---:|---:|---:|
| Numbers chart kind | 203,192,221 | 76,102,352 | −62.5% |
| Numbers chart direction | 203,128,434 | 76,094,695 | −62.5% |
| Numbers chart data | 217,889,592 | 105,218,075 | −51.7% |
| Pages chart kind | 1,921,555,557 | 49,756,268 | −97.4% |
| Pages chart direction | 1,921,878,214 | 49,770,449 | −97.4% |
| Pages chart data | 1,413,844,654 | 56,504,008 | −96.0% |

The Pages host re-saves the whole document through its logical writer, so
its routes cost far more.

## 5. Left

| Item | Rows | Next |
|---|---:|---|
| chart add, duplicate and remove (Numbers 3, Pages 3) | 6 | Lift Keynote's chart graph author (`slide_chart_lifecycle/create.rs`, the host's `charts::source`) into a shared engine, then author the chart under each leaf's own owner, stylesheet and theme-preset registration |
| Pages `set_body_chart_geometry` | 1 | A frame write must also move the body attachment, as the host does |
| shared `document.rs charts` | 1 | A cross-format aggregate: `litchi-iwa-structured` or the root coordinator |
| Numbers shape image fill | 1 | The data-owner transition on `structure::metadata::Transition` |
| Numbers shape, image and movie titles and captions | 15 | Lift the caption graph seam, then add a Numbers caption family |
| Pages `Builder::body_table` | 1 | A from-scratch table graph in `litchi_numbers_wire::table_create` |
| Keynote | — | Move its chart kind/direction and line-endpoint code onto the shared `chart_lifecycle` and `line_end` modules |

## 6. Verification

Run on `a18d806e4a` with this record's file in the working tree. The
build env follows the coordinator rules, and test executables were pruned
after each chunk. No litchi-ops dispatch arm changed. Every row exited 0:

- `cargo fmt --all --check`;
- strict clippy of litchi-iwa-style, -pages, -numbers, -iwa, -ops, litchi
  and litchi-py (all targets, all features);
- clippy of litchi-ops with no features, `numbers` and `pages`;
- `cargo test --all-features` of litchi-iwa-style, litchi-pages,
  litchi-numbers, litchi-iwa (parity included) and litchi;
- litchi-ops tests: `--lib`, `--doc` and nine `--test` groups;
- rustdoc `-Dwarnings`;
- `check_crate_boundaries.py` and `non_iwork_gate.py verify`;
- the glossary gate (`RUSTC_BOOTSTRAP=1`);
- pyo3 `run_tests.sh`;
- perf-baseline `cargo check --locked`;
- the tool unit tests (including `test_iwa_exit_ledger`);
- the ledger, capability-delta, corpus, bench-matrix and feature-matrix
  checks.

Both pre-commit size checks printed nothing for every commit.
