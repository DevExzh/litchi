# 0376c — W1-F Numbers formula authoring and cell-writer correctness

- Status: slice complete on `uo2/iwa-w1f`; coordinator trunk
  `6f27c738ad` is merged at `fbb40a9d05`, with route-attribution and
  scratch-probe cleanup at `32afc486de` and `72d414a13a`.
- Date: 2026-10-28.
- Scope: Numbers cell writes, formula dependency envelopes, formula-list key
  reuse, and `set_formula`/`set_formula_with_cached_value` parity.
- Native evidence: none. The migrated formula-authoring route is queued in
  `0363-mac-queue.md`; this record does not claim E3/E4.
- `performance_claim`: each measured leaf facade route is at least 55.74%
  lower in callgrind instructions than its corresponding host route on a
  separate pinned-CPU ABBA probe (§5).

## 1. What landed

| Commit | Area | Change |
|---|---|---|
| `0589134033` | `litchi-numbers-wire`, `litchi-numbers` | publish column-header cell counts for every stored-cell transition through the bounded batched header writer |
| `3fcd93bc58` | `litchi-numbers-wire` | preserve an existing empty `RangeDependenciesArchive` envelope, including an explicit empty envelope when the source omits the range payload |
| `17d18ba4c7` | `litchi-numbers-wire` | release zero-count formula-list roots before allocation and reuse the released key below `nextListID` |
| `a03370efdc` | `litchi-iwa`, `litchi-numbers`, ledger and queue | formula-author parity cases, typed bounded-header error mapping, ledger migration and native queue entry |
| `e912148f32` | branch integration | merge current `feat/unified-ops-api` (`7973aef0e6`) and resolve the parity module registry by unioning trunk's W2 modules with W1-F |
| `72d414a13a` | parity tests | remove all temporary diagnostic dump probes; retain only the documented cost probe |
| `32afc486de` | parity attribution | call the actual host `set_formula` for no-cache authoring, add direct cache-policy and reopened semantic projections, establish the initial no-cache `TST.Tile` normalization, and split the release cost probes by route; the normalization was narrowed by `e8aa7a16a7` |
| `e82bb18991` | branch integration | merge current trunk `0944b7319f` and retain the W1-G hidden-axes parity module |
| `e8aa7a16a7` | parity admission | narrow the no-cache normalizer to C3's known numeric display-cache bytes and add a neighboring-cell negative control |
| `bd787a55ce` | evidence correction | preserve the original unknown-field census, document the narrow normalizer/negative control and final per-route cost receipts, and record the scoped gates |
| `7477da37b1` | branch integration | merge coordinator trunk `aeac888d5f` after the scoped gates |
| `9c342c394b` | gate receipt | record Clippy, rustdoc, glossary, boundary, no-3rdparty, non-iWork and fuzz-lock results |
| `fbb40a9d05` | branch integration | merge coordinator trunk `6f27c738ad`; rerun focused parity and ledger validation |

## 2. Public route and correctness contract

The selector-first Numbers facade is:

```text
Package::edit_table_cells(sheet, table)
  -> Edit::set_formula(At, Expression)
  -> Edit::set_formula_cached(At, Expression, CachedValue)
  -> Commit::{package, patch, diagnostics}
```

The route retains the immutable source snapshot, exact-source patch and
`Patch::inverse()` contract. The parity route applies the inverse and confirms
the original bytes are restored; exact no-ops remain output-free. The formula
author facade is thin over the shared wire dependency/formula-list engines and
the existing bounded table writer.

Cell writes now collect one sorted column delta per stored-cell transition and
publish the column-header bucket on every write, including formula writes and
clears. `header_adjust_many` has a checked growth envelope of
`bucket.len() + deltas * HEADER_BYTES` and performs one bounded output
allocation.

The package boundary preserves typed failures from that writer:

| Wire condition | Public result |
|---|---|
| malformed or non-canonical bucket | `Error::InvalidSource { path }` |
| unsupported storage or unordered header dependency | `Error::Refused(UnsupportedDependency { kind: CellStorage, .. })` |
| input/output/field/item budget exhaustion | `Error::Limit` with `WireBytes`, `OutputBytes`, `WireFields` or `RetainedElements` |
| fallible output allocation | `Error::Allocation` with `RetainedBytes` |

No focused-route refusal falls back to the host. The regression test
`bounded_header_errors_keep_their_public_categories` pins the malformed,
limit and allocation mapping.

## 3. Host differential parity

Cases are in
`crates/litchi-iwa/tests/parity/numbers_formula_author_cases.rs` and are
registered by `parity/main.rs`.

| Test | Fixture and operation | Tier/evidence | Result |
|---|---|---|---|
| `numbers_table_formula_author_parity` | `basic.numbers`, C3 = B3 + 1; owner has no dependency tile at the origin | B / queued | pass |
| `numbers_table_formula_author_graph_parity` | `formula-budget-native.numbers`, C3 = B3 + 1 with explicit cache 2 beside the native formula graph | B / queued | pass |
| `numbers_table_formula_cached_parity` | `basic.numbers`, C3 = B3 + 1 with an explicit numeric cache 43 | B / queued | pass |
| `numbers_table_formula_author_host_set_formula_policy` | direct host `set_formula` versus leaf no-cache authoring, reopened formula/value and exact host-zero versus leaf-evaluated cache policy | focused regression | pass |
| `numbers_table_formula_author_cache_normalizer_negative_control` | mutate neighboring B3 outside the selected C3 display-cache bytes and require the canonical output to change | focused regression | pass |

The cases agree on known formula-owner, cell-record-tile, data-list and
header fields; untouched members/objects remain exact. Documented Tier-B
differences are deterministic leaf UUIDs versus host-random UUIDs, leaf root
preview invalidation, unknown protobuf fields, and ZIP/member-layout changes.
The no-cache host route now calls the actual host `set_formula`; it leaves a
zero display cache while the leaf evaluates and stores the authored scalar.
The Tier-B case normalizes only C3's known numeric display-cache bytes in the
modern BNC cell record. It preserves the selected row's offsets and every
other row/cell's storage, pre-BNC buffers, styles, formula keys, and opaque
tail bytes. `.read` projections compare the formula/value at C3, the
precedent at B3, and every materialized table-cell value. The focused policy
test asserts the exact reopened cache values (`0.0` host versus `43.0` leaf),
and the neighboring-cell negative control proves that an unrelated B3 mutation
is not hidden by the normalizer. The graph case uses explicit cache 2 because
the host no-cache setter intentionally leaves that fixture's native tile
untouched; it therefore exercises the host cached setter beside the native
formula graph rather than substituting a no-cache call.

The host's `formula-events-native.numbers` dependency shape remains refused by
the leaf's focused route and is not silently downgraded to a host fallback.
Formula replacement's host object-identifier policy differs from the leaf;
formula-list key reuse is covered by the wire unit test rather than claiming a
false byte/object parity case. Both are documented follow-ups.

## 4. Native verification boundary

The queue entry names
`litchi_numbers::Package::edit_table_cells::Edit::{set_formula,set_formula_cached}`,
the three parity cases above, and the `basic.numbers` and
`formula-budget-native.numbers` fixtures. It remains `pending` because no
macOS run has opened, saved, closed and reopened a Litchi-mutated output.
Synthetic round trips and host differential parity do not establish native
acceptance.

## 5. Performance

Prescribed release-mode callgrind measurement: `CARGO_BUILD_JOBS=4`,
`CARGO_PROFILE_RELEASE_DEBUG=1`, LTO-enabled release profile, CPU 3 pinned,
one ignored test in ABBA order (host, leaf, leaf, host) for each formula-
author route. Each probe measures the full host/leaf facade through candidate
serialization and excludes the parity harness inverse readback. The host
toggle calls the corresponding host API: direct `set_formula` for the basic
no-cache case and `set_formula_with_cached_value` for the two cached cases.

| Route | Host Ir | Leaf Ir | Delta |
|---|---:|---:|---:|
| Basic no-cache `set_formula` (`numbers_table_formula_author_cost_abba`) | 189,825,420 | 84,009,135 | **−55.744%** |
| Native-graph cached authoring (`numbers_table_formula_author_graph_cost_abba`) | 159,361,970 | 50,041,627 | **−68.599%** |
| Basic explicit cached authoring (`numbers_table_formula_cached_cost_abba`) | 189,856,594 | 83,989,918 | **−55.761%** |

Reproduction:

```sh
CARGO_BUILD_JOBS=4 CARGO_TARGET_DIR=target/w1f \
  CARGO_PROFILE_RELEASE_DEBUG=1 \
  cargo test --release -p litchi-iwa --test parity --no-run
taskset -c 3 valgrind --tool=callgrind \
  '--callgrind-out-file=(deleted-coord-artifact)' \
  '--toggle-collect=*numbers_formula_author_cases::host_route*' \
  target/w1f/release/deps/parity-bf69cf43b955a09c --ignored --exact \
  numbers_formula_author_cases::numbers_table_formula_author_cost_abba --test-threads=1
# Repeat the command for each test with `host_route`/`leaf_cost_route` and the
# corresponding output path:
#   numbers_table_formula_author_cost_abba       basic.host.narrow/basic.leaf.narrow
#   numbers_table_formula_author_graph_cost_abba  graph.host.narrow/graph.leaf.narrow
#   numbers_table_formula_cached_cost_abba        cached.host.narrow/cached.leaf.narrow
```

The non-committed raw captures are outside the checkout at
`(deleted-coord-artifact)`:

| Artifact | Size | SHA-256 |
|---|---:|---|
| `cg.w1f.formula-author.basic.host.narrow` | 724,655 bytes | `6af462e7ecd55d6b3a105e69525f526d547d5f2d878575f06a117677bf99a872` |
| `cg.w1f.formula-author.basic.leaf.narrow` | 1,169,748 bytes | `f9eea5216cbf4ed9d9bcb3c9e8bfbad0f90daf65ef80cf1486e5d6b8f0de655f` |
| `cg.w1f.formula-author.graph.host.narrow` | 666,346 bytes | `038e87cf20737e076e36fa6b8a3ecbaa990d1f53a1ed95a428d4ae01c98e310f` |
| `cg.w1f.formula-author.graph.leaf.narrow` | 1,108,361 bytes | `b662bf4fe2eba1f5a0fbb893d2299896e283b5d44c8b0370a7c53dc1d9e3b31f` |
| `cg.w1f.formula-author.cached.host.narrow` | 724,860 bytes | `99ec5323507f4bde96591f8f8b45148170fb87407578662ab78a07ad085635b` |
| `cg.w1f.formula-author.cached.leaf.narrow` | 1,170,958 bytes | `a065ea741dc14b93a6769eabaf4287461b36ca12378ae43588f5aea7b5c584e2` |

The measured binary is also outside the repository: source commit
`e8aa7a16a7`, `target/w1f/release/deps/parity-bf69cf43b955a09c`,
197,464,264 bytes, SHA-256
`e39f4751bbb010d8a916f264495f685476ae3e519e5c85e1ab654e1cfdd24a12`.

## 6. Verification on the merged candidate

The final WP candidate is `fbb40a9d05` plus this record-only handoff update.
The source-sensitive checks below were run on the same branch; the parity and
ledger checks were rerun after the final `6f27c738ad` merge, and the boundary
audit passed on the integrated WP candidate. Full integrated qualification
remains coordinator-owned.

| Command | Result |
|---|---|
| `cargo fmt --all --check` | pass |
| `cargo test -p litchi-numbers-wire --lib` | 191 passed |
| `cargo test -p litchi-numbers --lib table_cell_edit::tests` | 19 passed, 1 intentional ignored |
| `cargo test -p litchi-numbers --test table_cells` | 17 passed |
| `cargo test -p litchi-iwa --test parity numbers_formula_author_cases -- --nocapture` | 5 passed, 3 intentional cost probes ignored after `fbb40a9d05` |
| `cargo clippy -p litchi-numbers-wire -p litchi-numbers -p litchi-iwa --all-targets -- -D warnings` | pass |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps -p litchi-numbers-wire -p litchi-numbers -p litchi-iwa` | pass |
| `check_api_glossary.py --profile numbers-minimal` and `--profile numbers-full` | pass; 53 transaction families, 53 explicit result contracts |
| `python3 tools/iwa_exit_ledger.py --check` | pass after `fbb40a9d05` |
| `python3 -m unittest tools.test_no_3rdparty_references` | 4 tests passed |
| `python3 tools/check_crate_boundaries.py` | pass; 70 packages, 313 declarations, 12 explicit debt items |
| `python3 tools/non_iwork_gate.py verify` | pass; 47 bulk trees and 43 facade-safe trees |
| `cargo check --locked --offline --manifest-path crates/litchi-numbers-wire/fuzz/Cargo.toml --bins` | pass |
| `cargo check --locked --offline --manifest-path crates/litchi-numbers/fuzz/Cargo.toml --bins` | pass |

Full 53-row and post-merge gates remain coordinator-owned. The native queue,
final integrated performance qualification, and D7 unified-operations
dispatch follow-up remain before W3.
