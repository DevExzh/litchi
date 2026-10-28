# 0389 — W2-D: relocate source-built Numbers sheet duplication tests

- Status: implementation slice on `uo2/iwa-w2d-source-built-duplication`, cut
  from frozen coordinator source `9932c7d0691b2c14bbbd11a2dea7df41f66308b6`.
- Date: 2026-10-28.
- Scope: move the four source-built sheet/table duplication assertions to the
  focused Numbers test crate. This record is not a merge or W3 deletion
  approval.

## Scope and compatibility

| Before | Destination | Leaf route |
|---|---|---|
| `crates/litchi-iwa/tests/source_built_sheet_duplicate_tables.rs` | `crates/litchi-numbers/tests/source_built_sheet_duplicate_tables.rs` | `Builder`, `edit_table_lifecycle`, `edit_sheet_lifecycle`, selector-first cell writes/reads |

All four assertion families remain: single-table clone content and storage
independence, multi-table sheet clone independence, invalid-selector
transactionality, and same-name clone scoping to the selected source sheet.
Each retains exact source/inverse, ordering, semantic readback, and reopen
checks where applicable.

The host fixture's `FormulaExpression::Number` was only a numeric-cell seed;
the focused test uses the leaf selector-first numeric cell writer to preserve
the duplication/storage contract without introducing a host formula API or
changing the test's purpose.

## Verification

With `CARGO_BUILD_JOBS=2`, `CARGO_INCREMENTAL=0`, and
`CARGO_PROFILE_DEV_DEBUG=0`:

- `cargo test --offline -p litchi-numbers --test source_built_sheet_duplicate_tables -- --nocapture`: 4 passed.
- `cargo fmt --all -- --check`: pass.
- Scoped strict Clippy for the moved test: pass; existing vendor warnings are
  dependency warnings.
- `python3 -m unittest tools.test_no_3rdparty_references`: pass.
- No runtime, fixture, parity/oracle, ledger, native queue/count, or
  performance changes are included.
