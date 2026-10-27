# 0370j — W1-A slice 11: Numbers sheet duplicate

- Status: landed on branch `uo2/iwa-w1a`, after slice 10
  ([0370i](0370i-iwa-w1a-builder-options.md), merged on trunk).
- Date: 2026-10-27.
- `performance_claim`: the leaf sheet duplicate runs 92.9% below the host
  route (§4).

## 1. What landed

`litchi_numbers::sheet::lifecycle::Edit::duplicate(selector)` stages a
copy of a sheet. The copy is named `"<name>-1"` (then `"<name>-2"`, …), as
the host names it, and goes right after its source in both the sheet list
and the sidebar.

**What the copy gets:**

- The source's sheet record, renamed and without its drawables. The
  record's style and settings are kept.
- A populated copy of each table, at the same position and under the same
  name.

**How it is built:** in one structural working set, on the table-duplicate
core from record 0370h. That core is now
`table_lifecycle::duplicate::clone_table`, which places a copy on a given
sheet (`Placement`).

**What is refused:**

| Source | Refusal |
|---|---|
| A sheet holding drawables other than tables | `Refusal::SheetNotEmpty` |
| A table whose formulas read another table | `Refusal::UnsupportedDependency`, from the port of the host's `table_formula_graph_is_self_contained` |
| A form-based sheet | `Refusal::UnsupportedSource` |

**Host defects fixed** (mac-queue row):

- The copy had no sidebar node, and its tables stayed listed under the
  source sheet's node. The new `sidebar::{insert_node, move_node}` places
  the sheet's node after its source's and moves each table's node under
  it.
- A copied legacy table drawable (a `TST.TableInfoArchive` stored in the
  calculation engine's component) did not register its reference to its
  new parent sheet.

**Documented behaviour (as the host):**

- The copy shares the source's header, footer and guide storages.
- The root previews are dropped; the app regenerates them.

One ledger row is migrated: `duplicate_sheet` → `edit_sheet_lifecycle`.

## 2. Parity

| Test | Routes | Tier |
|---|---|---|
| `numbers_sheet_duplicate_parity` | `basic.numbers` (first sheet), and a two-sheet workbook whose added sheet holds a table (each sheet) | B |
| `numbers_sheet_duplicate_corpus_parity` | the first sheet of each corpus fixture: 14 routes; sheets with other drawables are the refusal above | B |

## 3. Gaps

| Item | Next step |
|---|---|
| sheets with charts, media, shapes or text boxes | the copy needs W2-A's drawable duplicate engines |
| tables whose formulas read another table | a sheet-wide remap of cross-table edges |
| shared header, footer and guide storages | deep copy, once a native copy is observed (mac queue) |

## 4. Performance

`cost_probes::numbers_sheet_duplicate_cost_abba` (the first sheet of
`basic.numbers`; host, leaf, leaf, host):

| Route | Host | Leaf | Δ |
|---|---:|---:|---:|
| sheet duplicate | 1,696,467,791 | 121,029,923 | **−92.9%** |

## 5. Verification

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-numbers-wire -p litchi-numbers -p litchi-iwa --all-targets --all-features -- -D warnings` | 0 |
| `cargo test -p litchi-iwa --all-features` | 0 |
| `cargo test --all-features -p litchi-numbers`, `-p litchi-numbers-wire` | 0 / 0 |
| `cargo test --all-features -p litchi-pages`, `-p litchi-keynote` | 0 / 0 |
| `cargo test -p litchi-ops --all-features --test numbers_unified` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-numbers` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile numbers-minimal` / `numbers-full` | 0 / 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |
