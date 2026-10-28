# 0376c — W1-F Numbers formula authoring and cell-writer correctness

- Status: slice complete on `uo2/iwa-w1f`; current trunk
  `7973aef0e6` is merged at `e912148f32`, with scratch-probe cleanup at
  `72d414a13a`.
- Date: 2026-10-28.
- Scope: Numbers cell writes, formula dependency envelopes, formula-list key
  reuse, and `set_formula`/`set_formula_with_cached_value` parity.
- Native evidence: none. The migrated formula-authoring route is queued in
  `0363-mac-queue.md`; this record does not claim E3/E4.
- `performance_claim`: the measured leaf facade route is 59.53% lower in
  callgrind instructions than the host route on the same three-case ABBA
  probe (§5).

## 1. What landed

| Commit | Area | Change |
|---|---|---|
| `0589134033` | `litchi-numbers-wire`, `litchi-numbers` | publish column-header cell counts for every stored-cell transition through the bounded batched header writer |
| `3fcd93bc58` | `litchi-numbers-wire` | preserve an existing empty `RangeDependenciesArchive` envelope, including an explicit empty envelope when the source omits the range payload |
| `17d18ba4c7` | `litchi-numbers-wire` | release zero-count formula-list roots before allocation and reuse the released key below `nextListID` |
| `a03370efdc` | `litchi-iwa`, `litchi-numbers`, ledger and queue | formula-author parity cases, typed bounded-header error mapping, ledger migration and native queue entry |
| `e912148f32` | branch integration | merge current `feat/unified-ops-api` (`7973aef0e6`) and resolve the parity module registry by unioning trunk's W2 modules with W1-F |
| `72d414a13a` | parity tests | remove all temporary diagnostic dump probes; retain only the documented cost probe |

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
| `numbers_table_formula_author_graph_parity` | `formula-budget-native.numbers`, C3 = B3 + 1 beside the native formula graph | B / queued | pass |
| `numbers_table_formula_cached_parity` | `basic.numbers`, C3 = B3 + 1 with an explicit numeric cache | B / queued | pass |

The cases agree on known formula-owner, cell-record-tile, data-list and
header fields; untouched members/objects remain exact. Documented Tier-B
differences are deterministic leaf UUIDs versus host-random UUIDs, leaf root
preview invalidation, unknown protobuf fields, and ZIP/member-layout changes.
For the no-cache case, the host differential supplies the host's evaluated
display cache deliberately: direct host `set_formula` leaves a zero display
cache while the leaf evaluates and stores the authored scalar. This isolates
formula/dependency payload parity and is recorded as an intentional cache
policy difference, not as native proof.

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
one ignored test in ABBA order (host, leaf, leaf, host) for each of the three
formula-author cases. The probe measures the full host/leaf facade through
candidate serialization and excludes the parity harness inverse readback.

| Route | Host Ir | Leaf Ir | Delta |
|---|---:|---:|---:|
| Numbers formula authoring, three-case ABBA collection | 539,307,965 | 218,237,275 | **−59.53%** |

Reproduction:

```sh
CARGO_BUILD_JOBS=4 CARGO_TARGET_DIR=target/w1f \
  CARGO_PROFILE_RELEASE_DEBUG=1 \
  cargo test --release -p litchi-iwa --test parity --no-run
taskset -c 3 valgrind --tool=callgrind \
  '--toggle-collect=*numbers_formula_author_cases::host_route*' \
  target/w1f/release/deps/parity-bf69cf43b955a09c --ignored --exact \
  numbers_formula_author_cases::numbers_table_formula_author_cost_abba \
  --test-threads=1
# repeat with leaf_cost_route as the toggle
```

The non-committed raw captures are outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/iwa-w1f/bench/`:

| Artifact | Size | SHA-256 |
|---|---:|---|
| `cg.w1f.formula-author.host` | 773,128 bytes | `d2d6ff72826bcc1bb37b40158a5ff2302bbce1f96f894f3b9da85d11d7094f9c` |
| `cg.w1f.formula-author.leaf` | 1,291,593 bytes | `6d4ec0ffbe41bfb7c36ef5fdc7a6d12d3bda5bea5f838c7f9321e6248b6340ff` |

The measured binary is also outside the repository: source commit
`72d414a13a`, `target/w1f/release/deps/parity-bf69cf43b955a09c`,
196,411,152 bytes, SHA-256
`e48b5f85a5b532aaa8ba68acd661feb2c3ee2e70bc3f6a22319c8ded742c11db`.

## 6. Verification on the merged candidate

| Command | Result |
|---|---|
| `cargo fmt --all --check` | pass |
| `cargo test -p litchi-numbers-wire --lib` | 191 passed |
| `cargo test -p litchi-numbers --lib table_cell_edit::tests` | 19 passed, 1 intentional ignored |
| `cargo test -p litchi-numbers --test table_cells` | 17 passed |
| `cargo test -p litchi-iwa --test parity numbers_formula_author_cases` | 3 passed, 1 intentional cost probe ignored |
| `python3 tools/iwa_exit_ledger.py --check` | pending coordinator rerun after integration |

Full 53-row and post-merge gates remain coordinator-owned. The native queue,
final integrated performance qualification, and D7 unified-operations
dispatch follow-up remain before W3.
