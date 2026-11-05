# 0405 — W2-D: relocate Numbers table lifecycle examples

- Status: source slice on `uo2/iwa-w2d-table-examples`, cut from frozen
  coordinator source `6adafb4f2d9a318276378a915fb973f30ba7723d`.
- Scope: relocate the four Numbers table lifecycle commands whose focused
  leaf routes are already present. No runtime, facade, parity, ledger, native
  queue, fixture, performance, or PROGRESS changes are included.
- `performance_claim`: none.

## Relocation and behavior

| Before | Destination | Focused route |
|---|---|---|
| `crates/litchi-iwa/examples/add_numbers_table.rs` | `crates/litchi-numbers/examples/add_numbers_table.rs` | `Package::edit_table_lifecycle(...).add(...)` |
| `crates/litchi-iwa/examples/duplicate_numbers_table.rs` | `crates/litchi-numbers/examples/duplicate_numbers_table.rs` | `Package::edit_table_lifecycle(...).duplicate(...)` |
| `crates/litchi-iwa/examples/resize_numbers_table.rs` | `crates/litchi-numbers/examples/resize_numbers_table.rs` | `Package::edit_table_resize(...).resize(...)` |
| `crates/litchi-iwa/examples/remove_numbers_table.rs` | `crates/litchi-numbers/examples/remove_numbers_table.rs` | `Package::edit_table_lifecycle(...).remove(...)` |

The add command preserves the historical sheet-index CLI. Duplicate, resize,
and remove preserve the historical workbook-wide table-index CLI by resolving
that index to a focused sheet selector plus sheet-local table selector. Each
command requires distinct input/output paths, uses the selector-first leaf
transaction, saves a distinct output, and reopens it for semantic verification.
Duplicate snapshots dimensions, materialized cells, and row/column headers;
add and resize verify dimensions; remove verifies the owning sheet's table
count. Every command applies the public inverse patch and checks exact source
bytes before saving output.

Only these four command surfaces relocate. Existing host differential/oracle
and native-provenance tests remain unchanged. The host examples are removed
only after equivalent focused leaf examples compile and run against the
tracked `test-data/iwork/numbers/basic.numbers` fixture. No native acceptance,
parity tier, or performance claim is made.

## Verification

Using offline mode, two build jobs, and no third-party inputs:

- `cargo check --offline -p litchi-numbers --examples` — pass.
- `cargo clippy --offline -p litchi-numbers --examples -- -D warnings` — pass;
  existing vendor warnings only.
- `cargo run --offline -p litchi-numbers --example add_numbers_table
  test-data/iwork/numbers/basic.numbers
  (deleted-coord-artifact)/add.numbers
  0 Added 3 4` —
  exit 0; output 137341 bytes, SHA-256
  `42664156bda43e9c767db342e9a1aa2ee28508ac68ac60c4174c1e4355147dc5`.
- `cargo run --offline -p litchi-numbers --example duplicate_numbers_table
  test-data/iwork/numbers/basic.numbers
  (deleted-coord-artifact)/duplicate.numbers
  0` —
  exit 0; copied 22×7 table with 2 materialized cells; output 137715 bytes,
  SHA-256 `e7b61483cd71bf5d77f744b6fdc95fbb94cb78941cd00e756a84ac22c0240f30`.
- `cargo run --offline -p litchi-numbers --example resize_numbers_table
  test-data/iwork/numbers/basic.numbers
  (deleted-coord-artifact)/resize.numbers
  0 24 8` —
  exit 0; 22×7→24×8; output 136382 bytes, SHA-256
  `6f7a0e1233db74c04f7a365284f5f486fd1c2130047224b1c1fad1ee8d9ffd9b`.
- `cargo run --offline -p litchi-numbers --example remove_numbers_table
  test-data/iwork/numbers/basic.numbers
  (deleted-coord-artifact)/remove.numbers
  0` —
  exit 0; 1→0 tables; output 126672 bytes, SHA-256
  `3a50da97186d533bf467770ae8d56cf147ce6687769bb69d1434b208aa543d16`.
- `cargo fmt --all -- --check` — pass.
- `cargo doc --offline -p litchi-numbers --no-deps` — pass.
- Numbers glossary `numbers-minimal` and `numbers-full` profiles — pass.
- `python3 tools/check_example_targets.py` — pass.
- `python3 tools/iwa_exit_ledger.py --check` — pass.
- `python3 tools/check_crate_boundaries.py` — pass: 70 workspace packages,
  314 internal declarations, 12 explicit debt items.
- `python3 -m unittest tools.test_no_3rdparty_references` — pass.
- Staged repository size and binary guards — no output.

No native acceptance, parity-tier, or performance claim is made by this
example relocation.
