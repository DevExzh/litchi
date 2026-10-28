# 0382 — D7: Numbers cached formula authoring through the unified cell route

- Status: bounded implementation slice on `uo2/iwa-d7-numbers-formula`, cut
  from frozen coordinator root `bb2475b8b0c11c8cef304178ee3fca1f590e9633`;
  source commit `8b3ef8b0fd585d4353633d9be29e5859d687d531` adds checked
  operation pre-admission before formula parsing and cache allocation.
  Coordinator integration and the full 53-row gate remain pending.
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
  to the existing leaf `Change::set_formula_cached` route. Before the
  property scan, formula parse, or cache construction, it checked-multiplies
  target count by property count and pre-admits that operation cost. The
  existing per-property bounds remain in place for diagnostics and atomicity.
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
- `cached_formula_repeated_casefolded_keys_keep_last_pair_and_reject_interleaved_value`
  verifies that repeated ASCII-case-insensitive `formula` and `cachedValue`
  assignments keep the last pair, while an interleaved `value` remains the
  native typed refusal and leaves the source unchanged.
- `cached_formula_operation_limit_precedes_parse_and_cache_allocations` sends
  a malformed 128 KiB formula and cache under a one-operation limit and checks
  for the typed `Operations` refusal (`observed=2`, `maximum=1`) before any
  parse/allocation work or mutation.

The direct-leaf comparison proves that `cachedValue` reaches the cached leaf
operation; formula text alone is not treated as cache evidence. Native
Numbers open/save/close/reopen evidence is absent. Existing queue row 241 in
[0363-mac-queue.md](0363-mac-queue.md) remains pending for
`litchi_numbers::Package::edit_table_cells::Edit::set_formula_cached`; this
slice adds no duplicate queue row.

## 3. Verification

| Check | Result |
|---|---|
| `cargo test -p litchi-ops --features numbers --test numbers_unified -- --nocapture` | 21 passed; 1 ignored |
| Numbers capability and registry golden tests with `LITCHI_BLESS=1` | passed; only the two Numbers goldens changed |
| `cargo fmt --all -- --check` | passed |
| `cargo clippy -p litchi-ops --features numbers --test numbers_unified -- -D warnings` | passed; existing vendor zlib warnings are outside the focused crate |
| Numbers glossary profiles, exit ledger, crate boundaries, no-`3rdparty` tests | passed |
| `cargo check --locked --offline --manifest-path crates/litchi-ops/fuzz/Cargo.toml --all-targets --features apply` | passed |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps -p litchi-ops --features numbers` | passed after removing the cfg-sensitive `Self::pages_options` links from the Numbers/Keynote option docs; existing vendor warnings are outside the documented package |
| Native evidence / queue | no evidence claimed; existing row 241 remains pending |
| Full unified-vs-direct-leaf instruction probe | passed on source `8b3ef8b0fd585d4353633d9be29e5859d687d531`; pinned `taskset -c 3` release callgrind, exact output equality, no wall timing |

### Branch-bound instruction receipt

The probe was built outside the repository with this complete command:

```text
CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/target \
CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=1 CARGO_BUILD_JOBS=4 \
CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 \
RUST_TEST_THREADS=1 PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH \
cargo test --release --locked --offline -p litchi-ops --features numbers --test numbers_unified --no-run
```

The source fixture is `test-data/iwork/numbers/formula-budget-native.numbers`,
101,345 bytes, SHA-256
`6726414adc52762fab5ce3c52d6244d8d55958cf9fc4ee8bb26c00234f048fd3`. The
release test binary is
`/home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/target/release/deps/numbers_unified-f84504c3e8179e98`,
45,449,512 bytes, SHA-256
`f67ab82972b9c4001350e24e36f4601b4ce0632a1036c5f30d87e92f94ffe079`.
The route symbols were checked with:

```text
nm -C /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/target/release/deps/numbers_unified-f84504c3e8179e98 | rg 'unified_cached_formula_route|direct_cached_formula_route'
```

Each pinned pass ran the ignored test with the exact filter and output file
below. Both logs report one passing test and exact unified/direct output
equality:

```text
taskset -c 3 valgrind --tool=callgrind --quiet --collect-atstart=no --toggle-collect='*numbers_unified::unified_cached_formula_route*' --callgrind-out-file=/home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/callgrind/unified.callgrind /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/target/release/deps/numbers_unified-f84504c3e8179e98 --ignored --exact cached_formula_full_file_cost_abba --test-threads=1 > /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/callgrind/unified.log 2>&1
taskset -c 3 valgrind --tool=callgrind --quiet --collect-atstart=no --toggle-collect='*numbers_unified::direct_cached_formula_route*' --callgrind-out-file=/home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/callgrind/direct.callgrind /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/target/release/deps/numbers_unified-f84504c3e8179e98 --ignored --exact cached_formula_full_file_cost_abba --test-threads=1 > /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/callgrind/direct.log 2>&1
callgrind_annotate --inclusive=yes --threshold=0 /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/callgrind/unified.callgrind > /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/callgrind/unified.annotation.txt
callgrind_annotate --inclusive=yes --threshold=0 /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/callgrind/direct.callgrind > /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/callgrind/direct.annotation.txt
sha256sum /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/target/release/deps/numbers_unified-f84504c3e8179e98 /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/callgrind/unified.callgrind /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/callgrind/direct.callgrind /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/callgrind/unified.log /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/callgrind/direct.log /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/callgrind/unified.annotation.txt /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/callgrind/direct.annotation.txt
stat -c '%s %n' /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/target/release/deps/numbers_unified-f84504c3e8179e98 /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/callgrind/unified.callgrind /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/callgrind/direct.callgrind /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/callgrind/unified.log /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/callgrind/unified.annotation.txt /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/callgrind/direct.log /home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/callgrind/direct.annotation.txt
```

The unified route recorded 75,491,093 Ir; the direct leaf route recorded
75,041,444 Ir, for +449,649 Ir (+0.599201%) on the unified route. The
annotation totals match those route totals. The raw receipts remain outside
the repository under
`/home/zhuhe/code/litchi-worktrees/.coord/d7-formula-perf-8b3ef8b/callgrind/`:

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| `unified.callgrind` | 1,198,882 | `672609d80cab23add3ada49b113bc1d9240471cfd084e90587233c813d87c49c` |
| `direct.callgrind` | 1,126,211 | `857b2b98a73fa68bb805a8fbf7f38aff625ce9af58bbb0090c2331191633ba23` |
| `unified.log` | 160 | `bb72c19c70891377e9dcdd2c38015bbf10701e482c8b5ad4f869440ab12bc6ea` |
| `direct.log` | 160 | `bb72c19c70891377e9dcdd2c38015bbf10701e482c8b5ad4f869440ab12bc6ea` |
| `unified.annotation.txt` | 1,190 | `8568ab2785f3825823ae8ff61ccb063d4f59fd3377d91dbedec2317db66f32ea` |
| `direct.annotation.txt` | 1,189 | `5e9411e117262595ef30a2a062736a71e0e8e23c5e5a1b2711c72750c30ee5bd` |

This is a source-bound branch receipt only; later merges require final
integrated requalification. It makes no native acceptance claim.

The coordinator must rerun the affected scoped gates and the complete 53-row
gate on the final integrated commit. This record does not authorize deleting
the host or parity harness.
