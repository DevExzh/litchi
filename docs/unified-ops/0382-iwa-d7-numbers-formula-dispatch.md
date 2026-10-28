# 0382 — D7: Numbers cached formula authoring through the unified cell route

- Status: bounded implementation slice on `uo2/iwa-d7-numbers-formula`, cut
  from frozen coordinator root `bb2475b8b0c11c8cef304178ee3fca1f590e9633`;
  follow-up source commit `796134552fc2061a10fa084e3f44d240dcd85cd2` closes
  the cfg-sensitive rustdoc link and adds the full-route probe. Coordinator
  integration and the full 53-row gate remain pending.
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
  inverse restoration, and refusal-without-mutation cases. Its ignored
  `cached_formula_full_file_cost_abba` test exercises the complete unified
  `File` open/set/publish route versus the direct leaf open/edit/commit/write
  route in `[unified, leaf, leaf, unified]` order and asserts exact bytes on
  every iteration.
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
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps -p litchi-ops --features numbers` | passed after removing the cfg-sensitive `Self::pages_options` links from the Numbers/Keynote option docs; existing vendor warnings are outside the documented package |
| Native evidence / queue | no evidence claimed; existing row 241 remains pending |
| Full unified-vs-direct-leaf instruction probe | passed on source `796134552fc2061a10fa084e3f44d240dcd85cd2`; pinned `taskset -c 3` release callgrind, exact output equality, no wall timing |

### Branch-bound instruction receipt

The probe was built outside the repository with:

```text
CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-7961345/target \
CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=1 CARGO_BUILD_JOBS=4 \
cargo test --release --locked --offline -p litchi-ops --features numbers \
  --test numbers_unified --no-run
```

The source fixture is `test-data/iwork/numbers/formula-budget-native.numbers`,
101,345 bytes, SHA-256
`6726414adc52762fab5ce3c52d6244d8d55958cf9fc4ee8bb26c00234f048fd3`. The
release test binary is
`numbers_unified-f84504c3e8179e98`, 45,344,912 bytes, SHA-256
`42deb0b45dd1f12ec31c79a22950f43e9fa7cf626588ba0991e3eeb6bb330bef`.
Each pass ran the ignored test with `--test-threads=1` and
`--collect-atstart=no`, toggling one route symbol at a time:

```text
taskset -c 3 valgrind --tool=callgrind --quiet --collect-atstart=no \
  --toggle-collect='*numbers_unified::unified_cached_formula_route*' ...
taskset -c 3 valgrind --tool=callgrind --quiet --collect-atstart=no \
  --toggle-collect='*numbers_unified::direct_cached_formula_route*' ...
```

The unified route recorded 75,482,316 Ir; the direct leaf route recorded
75,040,012 Ir, for +442,304 Ir (+0.589424%) on the unified route. The raw
receipts remain outside the repository under
`/home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-7961345/callgrind/`:

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| `unified.callgrind` | 1,198,150 | `b21365b6883ddab3d2094a0cf1ef78735d49da08691f993b2dc284b7d27c4d6c` |
| `direct.callgrind` | 1,125,965 | `d0ef9a2d02eeefb995009646ddcb3f6f5f71a5758cda44bb320d2a5bb3f3d932` |

This is a source-bound branch receipt only; later merges require final
integrated requalification. It makes no native acceptance claim.

The coordinator must rerun the affected scoped gates and the complete 53-row
gate on the final integrated commit. This record does not authorize deleting
the host or parity harness.
