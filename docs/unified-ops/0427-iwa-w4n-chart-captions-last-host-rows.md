# 0427 — W4-N: the componentized Numbers chart caption and Pages chart captions (the last host-only ledger rows)

- Status: merge-ready on branch `uo2/iwa-w4n` (cut from trunk
  `feat/unified-ops-api` at `27939b7955`).
- Date: 2026-11-03.
- Plan: [0363](0363-iwa-monolith-exit-plan.md) (§0, §3, §4 parity tiers,
  §8 D1); the predecessor's successor brief
  [0425](0425-iwa-w4l-series-image-fills-captions.md) §4; the shared
  caption engine of 0425 slice 2; owner decision D1 and the
  [macOS queue](0363-mac-queue.md).
- Scope (coordinator, wave W4-N): the four rows still host-only —
  Numbers `set_sheet_chart_caption` (the componentized layout) and Pages
  `body_chart_caption`, `set_body_chart_caption`,
  `remove_body_chart_caption`.
- Result: `python3 tools/iwa_exit_ledger.py --summary` reports **0
  host-only rows** (Keynote 634 migrated / 5 retired; Numbers 586 / 16;
  Pages 660 / 7; shared 237 / 91).
- `performance_claim`: every migrated route uses 62–94% fewer
  instructions than its host route under callgrind (§3); none comes near
  the 5% bar.

## 1. The shared engine: componentized layout and the chart caption edge

`4efdd0be87` (shared crates and Numbers, one commit because the
`graph_objects`/`write` signatures change with their callers).

| Crate | What |
|---|---|
| `litchi-iwa-protos` | `keynote_chart_caption_graph_codec::CaptionGraphProfile::Componentized` (the enum existed with `Inline` only, documented as the place for this variant). It authors the same style, info and storage payloads and swaps the placement anchors (caption: caption top / drawable bottom; title the reverse), as the host's `CaptionKind::placement(Componentized)`. Unit test `componentized_profile_swaps_only_the_placement_anchors` |
| `litchi-iwa-style` | `caption::Layout::{Inline, Componentized}` (the host's `CaptionGraphProfile`). `graph_objects` takes it: a componentized graph's caption info and placement carry `COMPONENTIZED_VERSION` ([10, 1, 0]) and the info's metadata the three `IgnoreAndPreserve` field infos (drawable text wrap 1.1.1.12, owned storage 1.6, drawable lock 1.1.1.13), as the host's `archive_object`. `write` takes it: a componentized new label puts its shape style into the document stylesheet (appended to `TSS.StylesheetArchive.styles`, referenced from the stylesheet's metadata, inserted in the stylesheet's member, UUID in the stylesheet's component, an external reference from the drawable's component — the host's `insert_componentized_caption_style`), the other three objects into the drawable's member with UUIDs there. `Nesting::Chart` and `retarget_references`: a chart message's metadata follows the host's chart caption edge (`retarget_*_chart_caption_edge`): the message lists the old label once and no field twice, the old label leaves each list and the new one is appended (images, movies and shapes keep `replace_reference`, in place). Unit tests for the edge and the componentized metadata |
| `litchi-keynote`, `litchi-pages` | pass `Layout::Inline` (no behaviour change) |
| `litchi-numbers` | `sheet::drawable::caption`: a chart caption is `Nesting::Chart`; a chart stored beside the calculation engine (`caption::component_holds(.., 4000)`, the host's `chart_caption_graph_layout`) writes in `Layout::Componentized`. `Refusal::ComponentizedLayout` is deleted (no alias) |

The first run of `numbers_sheet_chart_caption_set_parity` without the
chart edge showed the only difference from the host: the chart message
5021 listed the new caption info where the old stand-in was, the host at
the end. `Nesting::Chart` closes it; the 0425 removal case passed before
only because the host-created stand-in already sat last.

## 2. Pages chart captions

`7c10e53acc`.

| Crate | What |
|---|---|
| `litchi-pages` | `chart::caption::{Edit, Patch, Commit}` (errors, refusals, limit kinds and diagnostics re-exported from `drawable::caption`); `Package::{body_chart_caption, edit_body_chart_caption, apply_body_chart_caption}`, `Edit::{set, clear, before}`. The body-caption core (`body_drawable_caption.rs`) gains a private `Selection::{Drawable, Chart}`: a chart resolves through the chart families' rooted `BodyChartSelector` proof (`body_chart_arrangement::resolve_target`), reads the `TSD.DrawableArchive` at field 1 of the chart message and shares the label-slot validation; publication is one shared step (`publish`) for both families. New chart objects register their UUIDs in the chart's component (the host's `source.component_id`), drawables keep the document component. A chart's title stays the chart-title family's, so the other-label verification applies to drawables only. Host layout check: the host's Pages route always uses the inline graph (`caption_objects`), so no Pages componentized path exists |

The host's chart caption code and the Numbers host code are the same edge
transition; the Pages route uses `Nesting::Chart` too.

## 3. Parity, ledger, queue

| Host route | Destination | Parity | Tier |
|---|---|---|---|
| `numbers/editor/sheet_charts/caption.rs set_sheet_chart_caption` | `litchi_numbers::Package::edit_sheet_drawable_caption` | `numbers_sheet_chart_caption_set_parity` (a new componentized caption on the host-authored chart in `basic.numbers`, then text replacement on the host-labelled chart) | B |
| `pages/editor/charts/caption.rs body_chart_caption` | `litchi_pages::Package::body_chart_caption` | `pages_body_chart_caption_read_parity` (uncaptioned native chart, host-captioned chart; clearing an absent caption is an exact no-op) | A |
| `pages/editor/charts/caption.rs set_body_chart_caption` | `litchi_pages::Package::edit_body_chart_caption` | `pages_body_chart_caption_set_parity` (new inline graph on the native `chart-data-native.pages` chart, then replacement) | B |
| `pages/editor/charts/caption.rs remove_body_chart_caption` | `litchi_pages::Package::edit_body_chart_caption` | `pages_body_chart_caption_remove_parity` | B |

Tier B: the host draws random UUIDs for the new objects and the leaf
deterministic ones (`leaf_structural_metadata`); every other byte of every
changed object and member is equal, the inverse patch restores the input
bytes. `numbers_sheet_chart_caption_componentized_refuses` is replaced by
the set parity. No native fixture carries a chart caption and the host's
caption routes have no native receipt, so all three routes are
macOS-queued (three new rows; the W4-L Numbers row now points at the
W4-N row instead of the refusal). Leaf tests:
`crates/litchi-pages/tests/body_chart_caption.rs` (set, replace, clear,
inverse, typed refusals, exact no-op). Glossary rows: the shared caption
row (Layout, `Nesting::Chart`, `retarget_references`,
`COMPONENTIZED_VERSION`), the Numbers caption row, a new Pages
`chart::caption` row; `tools/api_glossary.json` registers the
`litchi_pages::chart::caption` family and `apply_body_chart_caption`.

Strict admission is unchanged: the leaf still validates the label graph
(caption-info shape, distinct objects, private storage and placement in
the drawable's component), refuses reserved markers, dependent storage
content, a missing width and non-exact sources, and the chart edge
refuses ambiguous chart metadata (old label listed twice, new label
already listed, data references).

## 4. Performance

Probes `numbers_caption_route_cost_abba` (`W4L_ROUTE`, two new routes)
and `pages_chart_caption_route_cost_abba` (`W4N_ROUTE`); release, CPU 3,
ABBA, inclusive Ir of the host and leaf route functions (open, one caption
operation, commit, serialize; two calls each).

| Route | Host Ir | Leaf Ir | Leaf/host |
|---|---:|---:|---:|
| Numbers chart_caption_create (componentized) | 349,426,883 | 133,027,390 | 0.3807 |
| Numbers chart_caption_remove (chart edge) | 283,406,844 | 84,125,663 | 0.2968 |
| Pages pages_chart_caption_create | 721,996,966 | 77,019,605 | 0.1067 |
| Pages pages_chart_caption_replace | 854,289,806 | 52,913,733 | 0.0619 |
| Pages pages_chart_caption_remove | 824,412,075 | 76,218,406 | 0.0925 |

The Numbers removal is unchanged from 0425 (84,107,402 there; +0.02%).
Receipt `results/0427/callgrind-chart-captions.json` (binary SHA-256,
build and run commands, per-capture size and SHA-256); raw captures in
`(deleted-coord-artifact)` (outside the repository). The bench matrix lists 0427
under `evidence_without_rows` (leaf-versus-host evidence).

## 5. Gates

Run on `7c10e53acc` + the record files (trunk `27939b7955`); logs
`(deleted-coord-artifact)` (one command at a time). Every command exits 0:

| Gate | Exit |
|---|---:|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-iwa-protos -p litchi-iwa-style -p litchi-numbers -p litchi-pages -p litchi-keynote -p litchi-iwa -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features numbers / pages / keynote] --all-targets -- -D warnings` (four runs) | 0 |
| `cargo clippy` of the five changed crates without features | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features` (five changed crates) | 0 |
| `cargo test --all-features` of `litchi-iwa-protos`, `litchi-iwa-style`, `litchi-numbers`, `litchi-pages`, `litchi-keynote`, `litchi-iwa` (parity included) | 0 |
| `cargo test -p litchi --all-features` | 0 |
| `cargo test -p litchi-ops --all-features` (`--lib`, `--doc`, nine `--test` groups of 20) | 0 |
| tool unit tests; `officecli_capability_delta.py --check`; `officecli_corpus.py --check`; `check_feature_matrix.py` (31 ok) | 0 |
| `unified_bench_matrix.py --check` | 1 in the run (the matrix did not list this record yet), 0 after regenerating it |
| `check_crate_boundaries.py`; `non_iwork_gate.py verify`; `iwa_exit_ledger.py --check`; `check_iwork_leaf_public_api.py` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (73 profiles, 778 families) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |

No dispatch, capability or corpus-visible behaviour changed, so no golden
was re-blessed. Target 23 GB after the gate run, 0.6 GB after pruning.
