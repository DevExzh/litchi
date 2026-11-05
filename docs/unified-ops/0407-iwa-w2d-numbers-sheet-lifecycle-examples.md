# 0407 — W2-D: relocate Numbers sheet lifecycle examples

- Status: source slice on `uo2/iwa-w2d-sheet-examples`, cut from coordinator
  source `7606f8315d05fba42086ba93ea316918cdb3aa07`.
- Scope: relocate the three Numbers sheet lifecycle commands whose focused
  leaf routes are already present. No runtime, facade, parity, ledger, native
  queue, fixture, performance, or PROGRESS changes are included.
- `performance_claim`: none.

## Relocation and behavior

| Before | Destination | Focused route |
|---|---|---|
| `crates/litchi-iwa/examples/add_numbers_sheet.rs` | `crates/litchi-numbers/examples/add_numbers_sheet.rs` | `Package::edit_sheet_lifecycle(...).add(...)` |
| `crates/litchi-iwa/examples/duplicate_numbers_sheet.rs` | `crates/litchi-numbers/examples/duplicate_numbers_sheet.rs` | `Package::edit_sheet_lifecycle(...).duplicate(...)` |
| `crates/litchi-iwa/examples/remove_numbers_sheet.rs` | `crates/litchi-numbers/examples/remove_numbers_sheet.rs` | `Package::edit_sheet_lifecycle(...).remove(...)` |

The commands preserve their historical input/output paths and zero-based sheet
index arguments. Each uses semantic sheet selectors, checks the committed
state, applies the public inverse patch against the exact source bytes, saves a
distinct output, and reopens it for semantic readback. Duplicate preserves the
source table count and the generated suffixed name. Removal rejects deletion
of the final sheet and preserves leaf errors for unsupported drawable graphs.

Only these three command surfaces relocate. Existing host differential/oracle
and native-provenance tests remain unchanged. No native acceptance, parity tier,
or performance claim is made by this example relocation.

## Verification

The tracked fixture is `test-data/iwork/numbers/basic.numbers` (136,357 bytes,
SHA-256 `f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`).
Using offline mode and the isolated target directory
`(deleted-coord-artifact)`, all three commands exited 0:

```text
cargo run --locked --offline -p litchi-numbers --example add_numbers_sheet -- \
  test-data/iwork/numbers/basic.numbers \
  (deleted-coord-artifact) Archive
cargo run --locked --offline -p litchi-numbers --example duplicate_numbers_sheet -- \
  test-data/iwork/numbers/basic.numbers \
  (deleted-coord-artifact) 0
cargo run --locked --offline -p litchi-numbers --example remove_numbers_sheet -- \
  (deleted-coord-artifact) \
  (deleted-coord-artifact) 1
```

The persisted outputs and command logs are retained outside the repository:

| Command | Result | Output SHA-256 |
|---|---|---|
| add | `created sheet "Archive": 2 sheet(s)`; 136,420 bytes | `421687751b2d2a2a06d459f6a4c23bbfe74df72b30bfa1775327f5ae0b48a480` |
| duplicate | `duplicated sheet "Sheet 1" as "Sheet 1-1" with 1 table(s)`; 80,836 bytes | `30bab8a69bcd5d32d0e72542a422486d12cb3664ba1b68959a34a6778b3be984` |
| remove | `removed sheet 1: 1 sheet(s) remain`; 79,401 bytes | `5e963eb411012ff066bd4be22ddc07e2c932e957722575bd582c745786cf763c` |

The commands also hash-checked their saved outputs after reopening and checked
the exact inverse patch against the in-memory source. Scoped checks passed:

- `cargo fmt --all -- --check`.
- `cargo check --locked --offline -p litchi-numbers --examples`.
- `cargo clippy --locked --offline -p litchi-numbers --examples -- -D warnings`.
- `python3 tools/check_example_targets.py`.
- `python3 tools/iwa_exit_ledger.py --check`.
- `python3 -m unittest tools.test_no_3rdparty_references`.
- `python3 tools/check_api_glossary.py --profile numbers-minimal` and
  `--profile numbers-full`.

The repository-wide boundary scan completed successfully from clean commit
`d38473dd21`: 70 workspace packages, 314 internal dependency declarations, and
12 explicit migration debt items passed. The external 604-byte log is
`(deleted-coord-artifact)/w2d-sheet-boundaries-d384.log`,
SHA-256
`137949bf1118a296ac903b5401ca287559b11d7e0bad05d3bc96591d5cc7a06c`.
The slice adds no dependency edges or boundary exceptions. No native
acceptance, parity tier, or performance claim is made.
