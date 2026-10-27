# 0379 — W1-G: the shared iWork table-family crate `litchi-iwa-table`

- Status: group 1 (cell formats) landed on branch `uo2/iwa-w1g`, cut from
  trunk `feat/unified-ops-api` at `ca202291ca`: the crate and the Numbers
  move (`35908bbc0a`, merge-ready at `7ea1e295b1`), then the Pages and
  Keynote families (this record's commit). Groups 2–5 (comments, rich text,
  topology, conditional highlights) are not started (§5).
- Date: 2026-10-26.
- Plan: [0363](0363-iwa-monolith-exit-plan.md) (W1 table families, §3.2
  "one family taking a format enum"), the W1-A Numbers table records
  [0370](0370-iwa-w1a-numbers-tables.md)–[0370c](0370c-iwa-w1a-row-column-removal.md),
  the W1-F formula engine [0376](0376-iwa-w1f-formula-cache-engine.md),
  the harness and ledger [0369](0369-iwa-exit-parity-harness-and-ledger.md).
- `performance_claim`: the Numbers format, control and Pop-Up Menu routes
  are unchanged under callgrind (+0.09%, −0.03%, +0.01%); the new Pages
  family runs 88.2% below its host route and the Keynote family 0.9% above
  its host route (§4).

## 1. Why a new crate

Numbers sheets, Pages body tables and Keynote slide tables store the same TST
graph: a `TableModelArchive`, its tiles of BNC cells and the `TableDataList`
sidecars. The leaf engines that edit cell formats, rich text, comments and
topology lived only in `litchi-numbers` (~62k lines of
`package/table_*.rs`), tied to its sheet selectors, its `Package` state and
its error types. W1-A had already moved payload pieces into
`litchi_numbers_wire`. The coordinator decided to put the format-neutral
engines into one focused crate that the three leaves consume privately and
no facade re-exports.

`litchi-iwa-table` engines take:

- a `TableSource` (lib.rs): the component catalog, the exact physical source,
  an object lookup, the source bytes and the ingress limits. Numbers
  implements it over its object index, Pages and Keynote over theirs;
- a resolved `cell_format::Target`: the table model's component, object and
  message route, its lock state, the cell, and two leaf-chosen coordinates
  (`scope`, `table`) that come back in errors;
- one `cell_format::TransactionBudget`.

They return private archive bytes per changed member and a neutral
`cell_format::Error { Refused, InvalidSource, Limit, Allocation,
Verification, PatchConflict }` with a content-free `Path`. The leaf keeps
selectors, reassembly policy, candidate reopen and its public types.

## 2. Group 1: cell formats

### 2.1 The move (`35908bbc0a`)

| From `litchi-numbers/src/package/` | To `litchi-iwa-table/src/cell_format/` |
|---|---|
| `table_cell_control_native.rs` | `control.rs` (git mv) |
| `table_cell_display_format_native.rs` | `display.rs` (git mv) |
| `table_cell_pop_up_menu_native.rs` | `pop_up_menu_native.rs` (git mv) |
| `table_cell_pop_up_menu_metadata.rs` | `pop_up_menu_metadata.rs` (git mv) |
| engine half of `table_cell_pop_up_menu.rs`: `TransactionBudget`, control and Pop-Up Menu reads, native rewrite, metadata transition | `budget.rs`, `pop_up_menu.rs` |
| `metadata.rs` route helpers used by the above | `metadata.rs` (copy; Numbers keeps its own for its other users) |

Numbers changes only at call sites: `package.rs` aliases the moved modules
under their old names (`pub(crate) use litchi_iwa_table::cell_format::control
as table_cell_control_native`, …); the Pop-Up Menu module converts the
neutral `Error`, `Path` and `LimitKind` to its public types with lossless
`From` impls (the neutral `scope` is the sheet position); the nine format
facades and the control facade map through the same conversion
(`native_path` now returns the engine path). The public API is unchanged.

### 2.2 Pages and Keynote

Plan §3.2 asks for one family taking a format enum instead of the host's
per-format families. Two engine modules make that possible:

- `cell_format::data_format::{read, rewrite, family, is_display_family,
  maximum_members}` dispatches the neutral `DataFormat` over the
  per-family readers and writers. `read` reads controls first, then the
  nine display families, and reports Automatic only when a family reader
  saw an automatic cell (a format no reader admits is
  `UnsupportedDependency`, never Automatic). `rewrite` changes one display
  family at a time;
- `cell_format::publish::reassemble` is the exact-source reassembly the
  Numbers display facades run after a rewrite (preflight bound, Snappy
  under the retained limits, ZIP splice), shared instead of copied a
  tenth time.

`TransactionBudget::with_list_keys(ListKeys)` selects how format and
control lists allocate keys. Numbers keeps `Reuse` (lowest free key, the
watermark stays on release), so its output is unchanged. Pages and Keynote
use `Append`: a new entry takes `next_list_id`, and a removed last key is
handed back, which is what the host's Pages/Keynote format writer and the
shared cell writer (`table_write::list_release`, record 0370a) do. The
parity harness found the difference (`next_list_id` 5 against 6 on a reset,
key 3 against 6 on a replacement); with `Append` both match the host.

New public families (glossary rows and registrations in `pages`,
`keynote-minimal`, `keynote-full`):

- `litchi_pages::table::format`: `Package::{body_table_cell_format,
  edit_body_table_cell_format, apply_body_table_cell_format}`,
  `Edit::{set(DataFormat), reset}`;
- `litchi_keynote::slide::table::format`: `Package::{slide_table_cell_format,
  edit_slide_table_cell_format, apply_slide_table_cell_format}`, the same
  verbs. The dedicated `edit_slide_table_cell_number_format` family stays.

Each resolves its selector (Pages `resolve_body_table`, Keynote
`slide_table_core::select_table`) into an engine `Target`, implements
`TableSource` (Pages scans its catalog, Keynote uses its object index),
reopens the candidate with its own `Package` and verifies the read. A
change between two families runs two engine rewrites on a private
intermediate candidate; the patch keeps the source and final artifacts,
so `inverse` restores the source byte for byte. Controls, Pop-Up Menus and
numeral systems are read, and refused for writing
(`Refusal::UnsupportedFormat`): their writes also rewrite the package
metadata registry, which needs the Numbers control facade's publication
(save tokens, preview policy) lifted too (§5).

Parity (`crates/litchi-iwa/tests/parity/{pages,keynote}_cell_format_cases.rs`),
on B2 of `body-table-cells-native.pages` and the first stored automatic cell
of `slide-table-read-native.key`:

| Test | Routes | Tier |
|---|---|---|
| `pages_body_table_cell_data_format_set_parity`, `keynote_slide_table_cell_data_format_set_parity` | Automatic to Number, host `set_table_cell_data_format` | B |
| `pages_body_table_cell_format_set_parity` | a replacement inside Number, Currency, Percentage, Scientific and Fraction on the host's formatted cell, host family setters | B |
| `keynote_slide_table_cell_format_set_parity` | the same for Currency, Percentage, Scientific and Fraction | B |
| `*_cell_format_reset_parity` | each of those families reset to Automatic, host family resets (Keynote Number: the generic setter) | B |
| `*_cell_format_read_parity` | host family and generic reads against the leaf read on the host's output | A |
| `*_cell_format_divergence_is_reported` | pins where the routes differ (§5) | — |

Ledger: 31 rows migrate, 17 Pages (`set_`, `reset_` and `table_cell_` for
the five decimal families, `set_table_cell_data_format`,
`table_cell_data_format`) and 14 Keynote (the same for the four families
without Number, plus the two generic rows). Six macOS queue rows cover the
routes; none has native evidence.

## 3. Tools and docs

- `tools/crate_boundaries.json`: the package with six edges (archive, core,
  protos, numbers-wire, common, core vocabulary) and `litchi-numbers ->
  litchi-iwa-table`; runtime-neutral.
- `tools/check_crate_boundaries.py`: the Numbers display-format, Custom,
  control and Pop-Up Menu source audits read the moved engine files
  (`IWA_TABLE_CELL_FORMAT_ENGINE_SOURCES`) for their markers and duplicate
  checks.
- `tools/non_iwork_gate.py` (+ test): 18 iWork packages, 21 excluded.
- `tools/api_glossary.json`: profile `iwa-table` (bans `_with_limits`
  suffixes and `LimitExceeded`).
- ADR 0024: topology amendment.

## 4. Performance

Callgrind, release with debug information (`lto = true`), CPU 3 pinned,
instruction counts. Probes (committed, ignored by default):
`tests/cell_format_cost_probe.rs` (number, currency, percentage, scientific,
fraction, duration, date-time and text set + clear on the data-format
fixture), `table_cell_control::control_cost_probe` (five controls set +
clear), `table_cell_pop_up_menu::pop_up_menu_cost_probe` (create, replace,
clear). Binaries of `59df2ec089` (before) and `35908bbc0a` (after) run
A, B, B, A.

| Route | Before (mean of A1, A2) | After (mean of B1, B2) | Δ |
|---|---:|---:|---:|
| display formats, 8 families | 32,233,052 | 32,263,295 | +0.09% |
| controls, 5 kinds | 14,444,461 | 14,439,824 | −0.03% |
| Pop-Up Menu | 4,736,325 | 4,736,939 | +0.01% |

Pages and Keynote, the same release binary,
`{pages,keynote}_cell_format_cases::*_cell_format_cost_abba` (a Currency
replacement on the host-written EUR cell; host family setter A, leaf
family B, run A B B A, open + edit + commit + serialize), collected with
`--toggle-collect=*probe_host*` and `*probe_leaf*`:

| Route | Host | Leaf | Δ |
|---|---:|---:|---:|
| Pages `set_table_cell_currency_format` → `edit_body_table_cell_format` | 3,361,490,406 | 396,996,392 | **−88.2%** |
| Keynote `set_slide_table_cell_currency_format` → `edit_slide_table_cell_format` | 785,765,067 | 792,722,687 | +0.9% |

The Keynote leaf route pays the slide-table selection proof and its
inventory; it stays inside the 5% bar.

## 5. Gaps

| Gap | Finding | Next step |
|---|---|---|
| Automatic to a decimal family other than Number, and cross-family changes on decimal cells | the engine admits a family only on a cell whose format entry already belongs to it, or an empty one; the host's compatibility writer converts (the Numbers 0370a §5 gap) | lift a conversion path into `data_format::rewrite` with parity against the host's `cell_data_format` writer |
| Date and Time, Duration, Text on number cells | the host writes a number under a date or duration kind, which the engine reads as an invalid cell; the host refuses Text on a number cell (pinned by `*_divergence_is_reported`) | decide with the macOS run which shape native files store |
| Custom, numeral system, controls, Pop-Up Menu for Pages and Keynote | read only; writes need the metadata transition and the control facade's publication | lift the Numbers control/Pop-Up Menu publication (`rewrite_transaction`, save tokens, preview policy) into the crate |
| empty cells | the engine edits stored cells only (`CellNotFound` on an empty slot); the host first stores the cell | reuse `table_write` to store a style-only cell first |
| groups 2–5 | cell comments (coordinator addition), cell rich text, topology orchestration and conditional highlights are not started | next W1-G slices, in that order |

## 6. Verification

(filled at the end)
