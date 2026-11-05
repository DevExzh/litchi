# 0389 — W2-D: focused source-built Numbers duplication coverage

- Status: implementation slice on `uo2/iwa-w2d-source-built-duplication`, cut
  from frozen coordinator source `9932c7d0691b2c14bbbd11a2dea7df41f66308b6`.
- Date: 2026-10-28.
- Scope: add three focused leaf duplication cases while retaining the original
  four-test host regression target unchanged. This record is not a merge or
  W3 deletion approval.

## Scope and compatibility

| Before | Destination | Leaf route |
|---|---|---|
| Host `crates/litchi-iwa/tests/source_built_sheet_duplicate_tables.rs` | unchanged | Host provenance regression target retained until W3 |
| Focused `crates/litchi-numbers/tests/source_built_sheet_duplicate_tables.rs` | new leaf target | `Builder`, `edit_table_lifecycle`, `edit_sheet_lifecycle`, selector-first formula/cached-value writes and reads |

The host target keeps all four assertion families unchanged: single-table
clone content and storage independence, multi-table sheet clone independence,
invalid-selector transactionality, and same-name clone scoping. The leaf target
adds three equivalent cases through focused lifecycle APIs: single-table
cloning, invalid-selector transactionality, and same-name scoping. The leaf
sheet duplication itself succeeds for the formula-bearing source; the later
formula edit on cloned `Alpha` fails before publication with typed
`FormulaWork { observed: 3, maximum: 1 }`. A formula-free numeric edit on the
same clone succeeds. The multi-table formula-bearing host regression therefore
remains host coverage; it was not replaced by a literal-cell test. The leaf
target adds exact inverse checks; the unchanged host target retains its
original readback/reopen assertions and makes no new inverse claim.

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
- Formula-owner diagnostic repro against frozen `9932c7d069`:
  `(deleted-coord-artifact)`;
  log SHA-256
  `1640144d7ff691b65c924ea54ee552255c83173bd412a5d9ef2a4d9387368c3b`;
  regenerate with `CARGO_BUILD_JOBS=2 CARGO_INCREMENTAL=0
  CARGO_PROFILE_DEV_DEBUG=0 cargo test --offline -p litchi-numbers
  --test formula_dup_repro -- --nocapture`.
- No runtime, fixture, parity/oracle, ledger, native queue/count, or
  performance changes are included.
