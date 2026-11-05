# 0370g — W1-A slice 8: sheet add/remove and table removal at host parity

- Status: landed on branch `uo2/iwa-w1a`, after slice 7
  ([0370f](0370f-iwa-w1a-table-add-alignment.md), merged on trunk).
- Date: 2026-10-26.
- `performance_claim`: the leaf lifecycle routes run 84.3% below the host routes (§4).

## 1. What changed

This slice closes the rest of record 0370e §3. Wherever the native app
and the host disagreed, the native behaviour wins and the host is fixed.

**Host defects fixed** (mac-queue rows):

- `add_empty_sheet`:
  - It wrote a bare `TN.SheetArchive` carrying only the name. A new sheet
    now takes the first sheet's settings (page setup, style, unknown
    fields) without the first sheet's private objects: drawables, the
    deprecated header and footer storages, the guide storage, headers and
    footers. The builder's own sheet carries these settings too.
  - It added no sidebar node. It now adds one under the root
    (`numbers::editor::sidebar`).
  - It left the package's last object identifier below the new sheet. The
    watermark is now raised.
  - The sheet's style reference is now registered across components.
- `remove_table`: it removed the drawable, the model, the cell storage and
  the formula owner, but left every other object that only the table
  addressed. Those objects include summary and category models with their
  own tiles, lists and headers, hidden-state, sort and pivot records, and
  table-private styles. The new `numbers::editor::object_closure` computes
  the table's private closure, and `remove_table` now removes all of it. A
  private object is one whose every referrer goes with the table, read from
  the archive headers' reference lists. This is the same rule as the
  focused crate's census.
- `remove_sheet`:
  - It left the sheet's sidebar node behind, still addressing the removed
    sheet.
  - It left the sheet's private objects behind. It now removes both.

**Leaf** (`litchi_numbers`):

- The sheet add, sheet remove and table remove routes keep the root
  previews, as every other W1-A family does.
- A removal that frees the top of the identifier space now lowers the
  watermark, as the host editor does:
  - The sheet and table removals call `Transition::release_suffix`.
  - When nothing else changes in the registry, `metadata::apply` lowers the
    watermark itself, because the addition batch only raises it.
- A sheet removal with tables is composed of several publications: each
  table goes first, then the sheet. It now advances the source's save token
  once (`Working::set_save_token`). Every component record that the
  composition advanced past the new token is brought back to it, so the
  whole removal counts as one save.

Three ledger rows are migrated:

- `add_empty_sheet` → `edit_sheet_lifecycle`;
- `remove_table` → `edit_table_lifecycle`;
- `remove_sheet` → `edit_sheet_lifecycle`.

## 2. Parity

`crates/litchi-iwa/tests/parity/numbers_lifecycle_cases.rs` replaces the
exploratory corpus test from 0370e:

| Test | Routes | Tier |
|---|---|---|
| `numbers_sheet_add_parity` | `basic.numbers` | A |
| `numbers_sheet_add_corpus_parity` | every corpus fixture | A |
| `numbers_table_remove_parity` | the first table of `basic.numbers` | B |
| `numbers_table_remove_corpus_parity` | the first table of each corpus fixture the host edits | B |
| `numbers_sheet_remove_parity` | `basic.numbers` with a leaf-added second sheet, empty or holding a 4×3 table, removing either sheet | B |
| `numbers_sheet_remove_corpus_parity` | the corpus's one multi-sheet fixture, which the leaf refuses (§5) | B |

The removals are tier B, not A. The host re-encodes the package metadata
and draws fresh object UUIDs where the leaf derives them; this is the
documented UUID policy.

## 3. Shared engine changes

- `structure::Working::{set_save_token, save_token}` and
  `metadata::save_token`.
- `metadata::apply` lowers the watermark on a bare released suffix.

## 4. Performance

`cost_probes::numbers_lifecycle_cost_abba` on `basic.numbers`: add a
sheet, remove the first table, and remove a second sheet that holds a 4×3
table. For each operation the probe runs host, leaf, leaf, host.

| Route | Host | Leaf | Δ |
|---|---:|---:|---:|
| sheet add + table remove + sheet remove | 3,139,369,674 | 492,947,539 | **−84.3%** |

## 5. Gaps

| Item | Next step |
|---|---|
| sheet removal when the sheet holds drawables other than tables (`table-relocation-native-resaved.numbers`): the leaf refuses, the host removes | uses W2-A's drawable deletion engines |
| table and sheet duplicate | next slice |
| no native evidence for the lifecycle routes | mac queue |

## 6. Verification

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-numbers-wire -p litchi-numbers -p litchi-iwa --all-targets --all-features -- -D warnings` | 0 |
| `cargo test -p litchi-iwa --all-features` | 0 |
| `cargo test --all-features -p litchi-numbers`, `-p litchi-numbers-wire` | 0 / 0 |
| `cargo test --all-features -p litchi-pages`, `-p litchi-keynote` | 0 / 0 |
| `cargo test -p litchi-ops --all-features --test numbers_unified` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-numbers` | 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |
