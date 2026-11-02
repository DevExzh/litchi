# 0412 — W4-T: shared cell rich-text engine, Pages and Keynote cell rich text

- Status: slices 1–3 merge-ready on branch `uo2/iwa-w4t`, cut from trunk
  `feat/unified-ops-api` at `4b26bb3379` and synchronized with trunk
  `1bc0d4c996`: engine lift `335baa93fc`, Pages `9eca49c6ae`, Keynote
  `2adfa1c782`, shared-publisher adoption `037c0132ec` (superseded by trunk's
  `list_workspace` in the second trunk merge), record `f8d34553fc`; merged
  with trunk `cd89e91f63`. Slice 4 (Pages/Keynote cell formats) is
  not in this record (§5).
- Date: 2026-11-02.
- Plan: [0363](0363-iwa-monolith-exit-plan.md) §3.2 (table-cell families),
  §4 (parity tiers); the Numbers families of
  [0370](0370-iwa-w1a-numbers-tables.md) (cell text properties) and
  [0370b](0370b-iwa-w1a-cell-paragraph-lists.md) (cell lists); the shared
  table crate of [0379](0379-iwa-w1g-shared-table-crate.md).
- `performance_claim`: every new and lifted route uses fewer instructions
  than its host route (§4).

## 1. Slice 1: the engine lift (`litchi_iwa_table::cell_rich_text`)

The Numbers cell text-property engine (`package/table_cell_text.rs`), the
cell-list engine (`package/table_cell_list.rs`) and the storage-list
publisher (`package/storage_list.rs`) move into `litchi-iwa-table` as
`cell_rich_text::{text, list, storage_list}`, generic over
`structure::TableLeaf` as `cell_appearance` and `hidden_axes` are. The
engines take a fresh `structure::Working` set and the table model's
identifier; each leaf resolves its own selector to that identifier and
keeps its public transaction, patch, verification and error vocabulary.

- `cell_rich_text` holds the shared graph helpers the two engines and the
  Numbers whole-cell text-style family use: table model geometry, tile cell
  location (`locate_cell`), style-list lookup, section base styles,
  `current_style`, stylesheet lookup, message replacement and the
  style-only/rich-text cell materializer (`materialize_cell`,
  `rewrite_cell`). The Numbers `table_cell_text_style` helpers of the same
  name now delegate to them; their unused originals are removed.
- `structure::wire` is the Numbers table-topology wire helper module, moved
  with `git mv` (Numbers re-exports it at its old path).
- The engine error is neutral (`Error { Refused(Refusal), InvalidSource,
  Limit { kind, observed, maximum }, Allocation, Verification }`, with
  `LimitKind { PackageBytes, PayloadBytes, Identifiers, Wire(_) }`); every
  facade converts it losslessly into its own error.
- Numbers keeps its public `cell::text` and `cell::list` families unchanged
  (thin facades). The list facade now locates the table model once per
  edit instead of twice per staged change.
- `litchi-iwa-table` gains the `litchi-iwa-text` dependency (the list value
  types); `tools/crate_boundaries.json` and the five affected locks (four
  fuzz workspaces and the FIN harness) are refreshed offline.
- Trunk's record 0413 lifted the same storage-list publisher into
  `litchi_iwa_text_wire::list::publish` behind a `Workspace` trait, with
  its working-set side in `litchi_iwa_table::structure::list_workspace`.
  This branch briefly carried its own adapter (`037c0132ec`); the trunk
  merge after it drops that module, and the cell-list engine publishes
  through trunk's `ListWorkspace`. No copy of the publication logic
  remains.

One shared text-engine change: `paragraph_style::plan` decodes a style's
own overrides only when the target variation is exclusive. Native Keynote
slide-table base styles carry a zero default tab interval, which the
override decoder refuses; a shared style is never rewritten, so the decode
was needless there, and the derived child does not copy the field. Numbers
behaviour is unchanged (all its parity and unit tests pass).

## 2. Slices 2 and 3: Pages and Keynote facades

| Family | Pages | Keynote |
|---|---|---|
| text properties | `litchi_pages::table::cell_text`: `Package::{body_table_cell_text, edit_body_table_cell_text, apply_body_table_cell_text}` | `litchi_keynote::slide::table::cell_text`: `Package::{slide_table_cell_text, edit_slide_table_cell_text, apply_slide_table_cell_text}` |
| lists | `litchi_pages::table::cell_list`: `Package::{edit_body_table_cell_lists, apply_body_table_cell_lists, body_table_cell_list, body_table_cell_lists, body_table_cell_list_levels, body_table_cell_list_level, body_table_cell_list_numbering, body_table_cell_list_property}` | `litchi_keynote::slide::table::cell_list`: the same names with `slide_table_` and a slide selector |

The verbs are the Numbers families' (`Edit::set(At, Property)`,
`Edit::reset(At, PropertyKind)`; `set_list`, `set_lists`, `set_level`,
`set_numbering`, `set`, `reset`). Selectors are semantic
(`BodyTableSelector`; `SlideSelector` + `TableSelector`); no native
identifier crosses the API. Each family is an immutable-owner transaction
(0351): `Commit { package, patch, diagnostics }`, an exact-source `Patch`
with `inverse()` and `is_noop()`, an exact no-op when every staged change is
skipped, typed `Refused`/`Limit`/`Allocation`/`Verification`/`PatchConflict`
errors. A Keynote selection that exhausts the selector budget reports
`LimitKind::Selection` instead of a flattened source error.

The 158 ledger rows (79 Pages, 79 Keynote) of `table.cell_rich_text`, plus
the six `paragraph_list_number_format` rows filed under `table.cell_format`
(they are list properties), migrate. The glossary registers the four
families in `pages`, `keynote-minimal` and `keynote-full`. Sixteen native
queue rows are appended (eight per format); no host route has native
evidence.

## 3. Parity

`crates/litchi-iwa/tests/parity/{pages,keynote}_cell_{text,list}_cases.rs`
reuse the Numbers cases' operations and samples on
`body-table-cells-native.pages` / `body-table-read-native.pages` (table 1)
and `slide-table-read-native.key` / `table-discovery.key` (slide 1, table
1):

| Test (`pages_body_table_` / `keynote_slide_table_` prefix) | Routes | Tier |
|---|---|---|
| `cell_text_property_parity` | 17 properties: set, reset on the host's output, then the next property (rewrite in place) | B |
| `cell_text_property_read_parity` | the 17 reads on five cells of every corpus table | A |
| `cell_list_preset_parity` | bullet/numbered presets and one-placement lists on plain cells (promotion), then reset, `None` and the other preset | B |
| `cell_list_level_numbering_parity` | level and numbering on plain and listed cells, level reset | A |
| `cell_list_property_set_parity` / `_rewrite_` / `_reset_` | the seven list properties | B / A / B |
| `cell_list_paragraph_parity` | a two-paragraph cell: mixed placements, level, numbering and properties of the second paragraph | A and B |
| `cell_list_refusal_parity` | a list property of a plain cell; a level at a non-start | both refuse |
| `cell_list_read_parity` | every list read on corpus cells and listed cells | A |

The Numbers cell text, text-style, list, text-box list, conditional
highlight and resize parity suites pass unchanged on the lifted engine.

## 4. Performance

Callgrind on the release parity binary of the final source (SHA-256
`06f14b1dd060d1fe86aec81f5451cbf6fec5270db46ac65ec20c4c05383fba4b`, built
with `cargo test -p litchi-iwa --all-features --test parity --release
--no-run`), CPU 5 pinned, `--toggle-collect` on the host and leaf wrappers
(`cost_probes::{host,leaf}_cell_{text,list}` for Numbers,
`pages_cost_probes::{probe_host,probe_leaf}` for the new probes), ABBA
order. Captures stay outside the repository in
`.coord/iwa-w4t/callgrind/s4-*.out`.

| Route (probe) | Host Ir | Leaf Ir | Leaf/host |
|---|---:|---:|---:|
| Numbers cell text (`cost_probes::numbers_cell_text_cost_abba`) | 626,523,340 | 323,175,816 | 0.516 |
| Numbers cell list (`cost_probes::numbers_cell_list_cost_abba`) | 1,715,000,464 | 274,259,411 | 0.160 |
| Pages cell text (`pages_cell_text_cases::pages_body_table_cell_text_cost_abba`) | 3,563,132,996 | 168,726,376 | 0.047 |
| Pages cell list (`pages_cell_list_cases::pages_body_table_cell_list_cost_abba`) | 10,626,996,029 | 167,361,884 | 0.016 |
| Keynote cell text (`keynote_cell_text_cases::keynote_slide_table_cell_text_cost_abba`) | 1,041,780,428 | 293,024,847 | 0.281 |
| Keynote cell list (`keynote_cell_list_cases::keynote_slide_table_cell_list_cost_abba`) | 2,218,871,473 | 291,258,149 | 0.131 |

The Numbers list leaf is at its pre-lift cost (record 0370b: 274,527,749):
a first lift build that located the table model twice per change measured
291,397,972 and was fixed before commit.

## 4a. Slice 4: Pages and Keynote cell formats

Engine (`litchi_iwa_table::cell_format`):

- `data_format::rewrite_family` routes a scalar control (Checkbox, Star
  Rating, Slider, Stepper) through the existing
  `control::rewrite_scalar_control` writer and every display family through
  `rewrite`; `maximum_members` admits the scalar writer's three members
  (tile, format list, control list). Pop-Up Menus stay refused
  (`Refusal::UnsupportedFormat` in the facades): they allocate a menu model
  and rewrite the package metadata registry, which needs the Numbers Pop-Up
  publication lifted (§5).
- **The 0370k gap is closed.** Native Pages and Keynote cells name an
  implicit Number entry (no explicit-format flag). A write of another
  decimal-kind display family (numeral system, Percentage, Scientific,
  Fraction) on such a cell no longer decodes that entry as the requested
  family: it interns a fresh canonical entry and releases the cell's
  reference to the implicit one, as the host's compatibility writer does.
  Explicit entries, Number on an implicit Number entry and every Numbers
  route behave as before (all Numbers format parity passes).

Facades: `edit_body_table_cell_format` / `edit_slide_table_cell_format`
now write the four scalar controls and numeral systems; the Pages and
Keynote format docs and refusal docs say so.

| Test (`pages_body_table_` / `keynote_slide_table_` prefix) | Routes | Tier |
|---|---|---|
| `cell_control_format_set_parity` | Automatic to Checkbox, Star Rating, Slider, Stepper, host family setters | B |
| `cell_control_format_reset_parity` | each control reset on the host's output, host family resets | B |
| `cell_control_format_read_parity` | host family reads against the leaf read on host-written controls | A |
| `cell_numeral_system_format_set_parity` / `_reset_parity` | Automatic to three numeral systems; their resets | B |
| `cell_data_format_set_parity` (extended) | Automatic to Number, Percentage, Scientific, Fraction | B |
| `pages_body_table_cell_text_format_read_parity` | host Text read against the leaf read on a host-written Text cell | A |

Ledger: 15 Pages and 14 Keynote `table.cell_format` rows migrate (set,
reset and read for the four controls; numeral set and reset; Pages Text
read). Five native queue rows are appended.

Callgrind (binary SHA-256
`062c9a6d7e564ef12e6bf62143dfeeffc286a7a3c9127d69eba3f38281a24752`, same
method as §4, `*_cell_{control,numeral}_format_cost_abba`):

| Route | Host Ir | Leaf Ir | Leaf/host |
|---|---:|---:|---:|
| Pages Automatic to Checkbox | 3,363,059,780 | 284,307,548 | 0.085 |
| Pages Automatic to numeral system | 3,358,583,373 | 398,422,531 | 0.119 |
| Keynote Automatic to Checkbox | 794,548,069 | 642,839,475 | 0.809 |
| Keynote Automatic to numeral system | 785,135,341 | 801,984,020 | 1.021 (within 5%) |

## 5. Gaps

| Item | Next step |
|---|---|
| shared rich-text payloads | refused, as on Numbers (0370b §5) |
| Pop-Up Menu writes and reads (Pages, Keynote; 3+3 rows) | the host refuses menus on the fixtures' number cells; leaf writes need the Numbers Pop-Up publication (menu model, metadata registry) lifted into `litchi-iwa-table` with a leaf trait for reassembly and reopen |
| Custom format (3+3 rows) | needs a document-scoped custom format in a fixture and the custom registry rewrite on Pages/Keynote |
| Date and Time, Duration (6+6 rows) | the host writes a number under a date/duration kind, which the engine reads as an invalid cell (0379 §5); a macOS run must decide the native shape |
| Text writes (2+2 rows) and Keynote Text read (1 row) | Text on a cell naming an implicit Number entry needs the converted-text metadata (secondary generic identifier); Keynote fixtures refuse the leaf/host cell writers (table formulas) |
| Automatic to Currency on an implicit Number entry | the cell's format kind changes too; still refused (pinned by the divergence cases) |
| no native rich-text cell in any corpus | macOS queue rows (§2) |

## 6. Verification

On the merge of trunk `1bc0d4c996` (sequential gate script
`.coord/iwa-w4t/gates.sh`, `CARGO_BUILD_JOBS=6`):

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-iwa-table -p litchi-iwa-text-wire -p litchi-numbers -p litchi-pages -p litchi-keynote -p litchi-iwa -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features numbers / pages / keynote] --all-targets -- -D warnings` | 0 each |
| `cargo test -p litchi-iwa-table -p litchi-iwa-text-wire -p litchi-numbers -p litchi-pages -p litchi-keynote --all-features` | 0 |
| `cargo test -p litchi --all-features` | 0 |
| `cargo test -p litchi-ops --all-features` (`--lib`, `--doc`, six chunks of 28 `--test` targets) | 0 each |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-iwa-table -p litchi-iwa-text-wire -p litchi-numbers -p litchi-pages -p litchi-keynote` | 0 (after the intra-doc link repair) |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile` `iwa-table`, `pages`, `keynote-minimal`, `keynote-full`, `numbers-minimal`, `numbers-full` | 0 each |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0 / 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_iwa_exit_ledger tools.test_check_iwork_leaf_public_api` | 0 |
| `python3 tools/officecli_capability_delta.py --check`; `officecli_corpus.py --check`; `unified_bench_matrix.py --check`; `check_feature_matrix.py` | 0 each (bench matrix regenerated for this record) |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `cargo check -q --locked --offline --manifest-path crates/{litchi-numbers,litchi-pages,litchi-keynote,litchi}/fuzz/Cargo.toml --bins`; `cargo metadata --locked --offline` on the FIN harness | 0 each |
| `cargo test -p litchi-iwa --all-features --test parity -- numbers_cell_ numbers_text_box numbers_sheet_text_box keynote_cell_ pages_cell_text_ numbers_conditional` | 0 (92 passed) |
| release parity binary: `pages_cell_list_cases keynote_cell_list_cases numbers_cell_list` | 0 (24 passed) |
