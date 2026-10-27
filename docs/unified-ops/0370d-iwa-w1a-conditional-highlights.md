# 0370d — W1-A slice 5: Numbers conditional highlights

- Status: landed on branch `uo2/iwa-w1a`, after slice 4
  ([0370c](0370c-iwa-w1a-row-column-removal.md), merged on trunk
  `d510b2e961`).
- Date: 2026-10-26.
- `performance_claim`: the leaf route runs 61.5% below the host route
  (§4).

## 1. What landed

New family `litchi_numbers::cell::conditional_highlight`:
`Package::{edit_table_cell_conditional_highlights,
apply_table_cell_conditional_highlights, table_cell_conditional_highlight}`
with `Edit::{set(At, rules), clear(At)}`. The value types (`Rule`,
`Condition`, `Style`, …) are re-exported from the archive-free
`litchi_iwa_common::table::cell::conditional_highlight` module. Four ledger
rows are migrated. The host's `cell_conditional_highlighting` returned
native identifiers with the rule count; the leaf read returns the rules,
so the count is their length.

The route ports the host editor's `conditional_highlight` module onto the
structural working set:

- `proto.rs` is a prost-free mirror of the schema subset the routes use
  (style set, predicates, arguments, formula AST nodes, paragraph and cell
  styles, colours, UUIDs). It encodes in field-number order as the
  generated encoder does. When it decodes a message whose emptiness a route
  checks, any field outside the subset counts as an override, so a native
  graph with unmodelled overrides is refused rather than misread.
- `native.rs`, `formula.rs` and `formula/**` are the host's predicate
  identifiers and canonical formula graphs, unchanged apart from types and
  error construction.
- `graph.rs` builds the style set, its paragraph and cell styles, the
  field info and the references. `read.rs` is the host's strict semantic
  decoder.
- `mod.rs` holds the three publications:
  - **create**: the list beside the table model, and the model's
    conditional-style owner, when the table has neither. The graph goes
    into the list's component, with UUIDs for every new object, then a list
    entry, then the cell's key and the applied rule.
  - **replace in place**: an entry the cell alone uses keeps its style
    identities, gains pairs for new rules and loses the surplus. A shared
    entry is released first.
  - **clear**: releases the entry and, when it went, removes the graph and
    its registry records and lowers the watermark.

  The routes stamp the same components' save tokens as the host. A new
  `Transition::stamp` does this, through the metadata codec's
  addition/save-token and save-token-only batches.

Volatile date rules (today, yesterday, tomorrow, the next or last period,
an offset from today) need the calculation engine's twelve-owner volatile
family, which the host creates with a random global-name UUID. The leaf
refuses them (`Refusal::VolatileRule`), and it refuses a cell that already
hosts one.

**Host defect fixed** (mac-queue row): a cell stored for its rules (an empty
cell) was counted in its row header but not its column header.
`update_cell` now goes through `update_tile_and_headers`, as every other
cell write does.

## 2. Parity

`crates/litchi-iwa/tests/parity/numbers_conditional_highlight_cases.rs`, on
B2, C3 and E7 (text, number and empty cells) of `basic.numbers` and
`custom-number-native.numbers`:

| Test | Routes | Tier |
|---|---|---|
| `numbers_conditional_highlight_set_parity` | 28 non-volatile condition kinds × 3 styles (fill, text colour, all three) as single rules: 168 | B |
| `numbers_conditional_highlight_replace_parity` | a three-rule set; then one, five and three rules in place; a second cell on the same list | B |
| `numbers_conditional_highlight_clear_parity` | clear after a three-rule set | B |
| `numbers_conditional_highlight_read_parity` | rules and count on corpus cells and on seven written rule sets per fixture | A |
| `numbers_conditional_highlight_volatile_refusal` | a `DateIsToday` rule is refused | — |

Tier B covers the fresh object UUIDs and the model's fresh conditional-style
owner UUID, which the host draws at random. The test's table-model decoder
drops that one field, and the case documents this as a policy.

## 3. Shared engine changes

- `structure::Working::upsert` replaces an object in place in its component.
- `structure::metadata::Transition::stamp`, with `apply` choosing the
  save-token-carrying batch whenever a transition stamps.

## 4. Performance

`cost_probes::numbers_conditional_highlight_cost_abba`: two rules on B2 of
`basic.numbers`, then a clear on the host's output; host, leaf, leaf, host.

| Route | Host | Leaf | Δ |
|---|---:|---:|---:|
| conditional highlight set + clear | 507,204,473 | 195,084,655 | **−61.5%** |

## 5. Gaps

| Item | Next step |
|---|---|
| volatile date rules | port the host's `dependencies.rs` owner-family registration (shared with W1-F's formula-owner engine), with a deterministic global-name UUID and a documented policy |
| no native conditional-highlight fixture | mac queue |

## 6. Verification

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-numbers -p litchi-iwa --all-targets --all-features -- -D warnings` | 0 |
| `cargo test -p litchi-iwa --all-features` | 0 |
| `cargo test -p litchi-numbers --all-features` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-numbers` | 0 |
| `cargo test -p litchi-ops --all-features --test numbers_unified` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile numbers-minimal` / `numbers-full` | 0 / 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |
