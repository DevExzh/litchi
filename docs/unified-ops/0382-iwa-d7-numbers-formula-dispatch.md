# 0382 — D7: Numbers cached formula authoring through the unified cell route

- Status: bounded implementation slice on `uo2/iwa-d7-numbers-formula`, cut
  from frozen coordinator root `bb2475b8b0c11c8cef304178ee3fca1f590e9633`;
  coordinator integration and the full 53-row gate remain pending.
- Date: 2026-10-28.
- Plan: [0363](0363-iwa-monolith-exit-plan.md) §0 D7; parity and evidence
  policy [0369](0369-iwa-exit-parity-harness-and-ledger.md).
- Scope: expose the existing Numbers leaf formula-cache author through the
  existing selector-first `set cell` dispatch. This slice does not add a
  verb, compatibility alias, leaf API, ledger migration, or native claim.

## 1. Public contract

The route accepts the existing `formula` property and the new set-only
`cachedValue` property together:

```text
set /sheet[1]/table[1]/C3 formula="=B3+1" cachedValue=2.0
```

`cachedValue` is declared as a scalar with `set=true`, `get=false`, and
`add=false` in both the Numbers capability and registry schemas. The unified
tree and its existing reader intentionally do not claim cache readback;
reopened formula/value projections remain the supported read shape.

The cache accepts finite numbers, booleans, and bounded text. Dates and
durations are not flattened into a scalar. A cache without a non-null formula,
a null cache, an unsupported scalar, or a non-finite number is refused with a
typed terminal error before any cell change is staged. Formula-only writes and
`formula=null` retain their prior semantics. Formula and cache are parsed and
validated as one assignment, while unrelated property order is preserved.

## 2. Implementation and proof

- `crates/litchi-ops/src/format/iwork/numbers.rs` maps the combined assignment
  to the existing leaf `Change::set_formula_cached` route and keeps formula
  parsing and cache construction before staging.
- `crates/litchi-ops/src/format/iwork/numbers/coalesce.rs` recognizes the
  combined cell write without changing the formula-only coalescer behavior.
- `crates/litchi-ops/tests/numbers_unified.rs` compares the unified output byte
  for byte with a direct leaf `set_formula_cached` commit from the same source
  and values. It separately compares formula-only output with the direct
  uncached leaf route, then checks reopened formula projection, exact no-op,
  inverse restoration, and refusal-without-mutation cases.
- The Numbers capability and registry goldens record the set-only schema.

The direct-leaf comparison proves that `cachedValue` reaches the cached leaf
operation; formula text alone is not treated as cache evidence. Native
Numbers open/save/close/reopen evidence is absent. Existing queue row 241 in
[0363-mac-queue.md](0363-mac-queue.md) remains pending for
`litchi_numbers::Package::edit_table_cells::Edit::set_formula_cached`; this
slice adds no duplicate queue row.

## 3. Verification

| Check | Result |
|---|---|
| `cargo test -p litchi-ops --features numbers --test numbers_unified -- --nocapture` | 19 passed |
| Numbers capability and registry golden tests with `LITCHI_BLESS=1` | passed; only the two Numbers goldens changed |
| `cargo fmt --all -- --check` | passed |
| `cargo clippy -p litchi-ops --features numbers --test numbers_unified -- -D warnings` | passed; existing vendor zlib warnings are outside the focused crate |
| Numbers glossary profiles, exit ledger, crate boundaries, no-`3rdparty` tests | passed |
| `cargo check --locked --offline --manifest-path crates/litchi-ops/fuzz/Cargo.toml --all-targets --features apply` | passed |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps -p litchi-ops --features numbers` | blocked by the pre-existing `options.rs:347` link to missing `Self::pages_options`; no owned docs changed |
| Native evidence / queue | no evidence claimed; existing row 241 remains pending |
| Performance | no wall-time claim; coordinator final qualification required |

The coordinator must rerun the affected scoped gates and the complete 53-row
gate on the final integrated commit. This record does not authorize deleting
the host or parity harness.
