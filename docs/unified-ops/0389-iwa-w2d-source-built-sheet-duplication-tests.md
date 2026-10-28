# 0389 — W2-D: relocate source-built Numbers sheet duplication tests

- Status: implementation slice on `uo2/iwa-w2d-source-built-duplication`, cut
  from frozen coordinator source `9932c7d0691b2c14bbbd11a2dea7df41f66308b6`.
- Date: 2026-10-28.
- Scope: add focused leaf coverage for the three source-built duplication
  assertions supported by the current lifecycle APIs, while retaining the
  original four-test host regression target unchanged. This record is not a
  merge or W3 deletion approval.

## Scope and compatibility

| Before | Destination | Leaf route |
|---|---|---|
| Host `crates/litchi-iwa/tests/source_built_sheet_duplicate_tables.rs` | unchanged | Host provenance regression target retained until W3 |
| Focused `crates/litchi-numbers/tests/source_built_sheet_duplicate_tables.rs` | new leaf target | `Builder`, `edit_table_lifecycle`, `edit_sheet_lifecycle`, selector-first formula/cached-value writes and reads |

The host target keeps all four assertion families unchanged: single-table
clone content and storage independence, multi-table sheet clone independence,
invalid-selector transactionality, and same-name clone scoping. The leaf
target relocates the three equivalent families that currently pass through
focused lifecycle APIs: single-table cloning, invalid selector
transactionality, and same-name scoping. The multi-table formula-bearing sheet
clone remains host-only because the current leaf transaction refuses that
source with a typed `FormulaWork` limit; it is not weakened to a literal-cell
test. The leaf target adds exact inverse checks; the unchanged host target
retains its original readback/reopen assertions and makes no new inverse claim.
Both targets preserve ordering and semantic checks where supported.

The leaf tests use `Expression::number` with `CachedValue::number`, preserving
formula-plus-cache semantics for the supported duplication families. The
multi-table formula-bearing host regression remains the authoritative host
coverage until the leaf transaction's formula-work refusal is resolved.

## Verification

With `CARGO_BUILD_JOBS=2`, `CARGO_INCREMENTAL=0`, and
`CARGO_PROFILE_DEV_DEBUG=0`:

- `cargo test --offline -p litchi-numbers --test source_built_sheet_duplicate_tables -- --nocapture`: 3 passed.
- Original host target remains present and unchanged for its four-test
  formula-bearing regression family.
- `cargo fmt --all -- --check`: pass.
- Scoped strict Clippy for the moved test: pass; existing vendor warnings are
  dependency warnings.
- `python3 -m unittest tools.test_no_3rdparty_references`: pass.
- No runtime, fixture, parity/oracle, ledger, native queue/count, or
  performance changes are included.
