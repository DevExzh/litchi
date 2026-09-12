# 0536 CFB cold-error-helper source review

`scope: read-only design and contract review for the 0536 matched CFB collector experiment`

`reviewed revision: 8876e87b8dbcb7dca54a3416386bbfc892c68eb4`

`reviewed baseline source: crates/litchi-cfb/src/file.rs SHA-256
72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b`

`reviewed candidate patch: docs/performance/results/change-0536/candidate.patch
SHA-256 29f96e9a19d8eb34508b84a84d1ac535bc970ace3e91d78f41d9a791ec05d9ec`

`reviewed formatted candidate source: crates/litchi-cfb/src/file.rs SHA-256
d8e322fb598950f157e9d290fbaf57567d0d7b3d9d5ec2c8d6a943becfe86a02`

`candidate diff review: pass; one production file (multiple hunks), no bitset
or test/harness changes; formatting-only delta from the reviewed draft`

`performance claim: none`

This review covers the frozen [0536 plan](plan.json), the private
[`CheckedBitSet`](../../../../crates/litchi-cfb/src/file.rs#L22),
[`SectorChainScratch`](../../../../crates/litchi-cfb/src/file.rs#L2764),
`collect_exact` and its stream-validation callers. The initial read-only pass
preceded patch materialization; the materialized patch was subsequently
reviewed below. No Rust source, test, build, benchmark, profiler, or capture
was run or changed for this review.

## Candidate boundary

The materialized patch adds eight private associated helpers to
`SectorChainScratch`, each marked `#[cold] #[inline(never)]`:
`chain_error_empty`, `chain_error_invalid_start`, `chain_error_length`,
`chain_error_index`, `chain_error_cycle`, `chain_error_exceeds`,
`chain_error_early_end`, and `chain_error_marker`. It replaces the eight
formatted diagnostics in `collect_exact` with calls to those helpers. The
`usize` conversion, explicit table bound, and defensive table-`get` paths all
use `chain_error_index`, so they retain one exact message. The patch does not
modify `CheckedBitSet`, `prepare_visited`, reservations, the loop's success
operations, or any stream-validation caller.

This is the intended narrow experiment: rare formatted diagnostics move out
of the ordinary `SectorChainScratch::collect_exact` path into private cold
helpers. The helpers receive only a borrowed `table_name` and copied scalar
values needed to produce the existing message. They add no successful-path
allocation or formatter call. The inspected source diff retains the
successful loop, checked accesses, marker checks, reservations, and reset
points semantically unchanged.

The 0535 instruction map is the relevant diagnostic boundary. Its valid-loop
range accounts for 99.9325% of timed XLS-owned collector self Ir and 99.9723%
of timed CFB collector self Ir, measured per workload in the current
[mechanism review](../change-0535/mechanism-review.md). That is an instruction
location result, not evidence that error formatting causes native latency.
The same map found no positive `collect_exact` call edge to
`CheckedBitSet::insert`; the collector emits its checked word test/set and
sector append directly. Consequently, moving the out-of-line bitset error
body alone has no demonstrated effect on this collector's hot loop. If the
candidate touches `CheckedBitSet::insert` for other callers, its direct
contract still applies below and the measured binary must show which callers
actually changed.

The accepted 0533 `claim_sector` cold helpers are outside this experiment's
candidate boundary. Do not fold physical-sector claiming, reconciliation,
provider freshness, or any earlier visited-bit representation into this
layout probe.

## Checked bitset obligations

`CheckedBitSet` is private, but its behavior is visible through chain,
directory, and mini-sector validation. A code-layout change must retain all
of the following:

| Operation or state | Required behavior |
| --- | --- |
| `try_with_capacity(bit_len, resource)` | Compute `bit_len.div_ceil(64)`, reserve words fallibly, map failure through the supplied resource label, then zero-initialize the words. |
| `contains(bit)` with `bit >= bit_len` | Return `false` without an error or allocation. |
| `contains(bit)` within `bit_len` | Derive the same word and mask and use a checked word read; a retained stale bit must not become visible after reuse. |
| `insert(bit)` with `bit >= bit_len` | Return `OleError::CorruptedFile` with exactly `bit index {bit} exceeds checked bit-set capacity {bit_len}`. |
| `insert(bit)` with no backing word | Return `OleError::CorruptedFile` with exactly `bit index {bit} has no backing word`; this is a defensive invariant path. |
| successful `insert(bit)` | Set only the requested bit, return `Ok(())`, and perform no allocation. Repeating an insert is idempotent. |

The helper must not turn a checked `get_mut` into indexing, make an invalid
bit panic, or alter the fact that a failed insertion leaves the words
unchanged. Error formatting belongs only on the error path; no successful
`collect_exact` iteration may acquire a new formatter call.

## Collector error order and exact messages

`collect_exact` resets the retained result at entry and resets it again when
its validation closure returns an error. The required order is:

1. Clear the previous sector result and set the visited map's logical length
   to zero.
2. For `expected_count == 0`, accept only `start_sector == ENDOFCHAIN` and
   return an empty result. Otherwise report the empty-chain error before any
   other preflight or allocation.
3. For a nonempty declaration, reject `start_sector >= MAXREGSECT`, then
   reject `expected_count > allocation_table.len()`.
4. Reserve `sectors` first with resource `sector-chain entries`; grow and
   clear the visited map next with resource `sector-chain map`.
5. For each declared entry, safely convert the sector to `usize`, check the
   allocation-table bound, check `contains`, call `insert`, append the sector
   in chain order, and load the next marker through checked `get`.
6. On the final entry require `ENDOFCHAIN`. On an intermediate entry reject
   `ENDOFCHAIN`, reject any marker at or above `MAXREGSECT`, and continue with
   the next sector.
7. On every error clear the sectors and set `visited.bit_len` to zero while
   retaining the allocated vectors. On success expose exactly the declared
   sector sequence and leave `visited.bit_len == allocation_table.len()` for
   a nonempty chain, or zero for an empty chain.

These exact diagnostics are part of the compatibility contract. Helpers may
centralize their construction but may not change capitalization, punctuation,
field values, or which malformed condition wins:

```text
Empty {table_name} chain must start with ENDOFCHAIN
Invalid start marker for {table_name} chain
{table_name} chain length exceeds its allocation table
Invalid sector index {sector} in {table_name}
Cycle detected in {table_name} chain at sector {sector}
{table_name} chain exceeds its declared length
{table_name} chain ends before its declared length
Invalid sector marker 0x{next:08X} in {table_name} chain
```

The invalid-index text is shared by the checked conversion, explicit bounds,
and defensive table-`get` paths. The intermediate marker checks occur only
after the current sector has been validated and recorded. Moving a helper call
earlier can therefore change the first error for a malformed table even when
the final valid result is unchanged.

## Allocation, reset, and ownership boundaries

All possible successful-path growth in `collect_exact` occurs before its loop:
the exact sector reservation precedes visited-map growth and zeroing. A map
growth failure retains the `sector-chain map` resource; a sector-vector
failure retains `sector-chain entries`. A cold diagnostic helper may allocate
only as part of constructing its error, as the current `format!` calls do; it
must not add an allocation or fallible operation to successful walking. No
candidate may replace `try_reserve_exact` with an infallible reserve, reserve
inside the loop, or change resource labels/order.

`prepare_visited` may retain more words than the current table. It sets the
new logical length and fills every retained word with zero on each nonempty
collection. `reset` clears only the sector length and logical bit length, so
capacity and backing-word pointers remain reusable after success and error.
The candidate must preserve this freshness boundary and must not expose stale
bits when a later collection uses a smaller or larger table.

`validate_stream_allocations` intentionally owns separate scratch values for
MiniFAT and regular FAT. They have independent table lengths and namespaces.
Each stream must be completely collected before `claim_chain` publishes
physical ownership; a failed collection must not claim a prefix. The
candidate must leave `claim_chain`, physical-sector roles/reconciliation,
root-chain handling, and MiniFAT/regular-FAT separation untouched.

The source also retains checked `u32` to `usize` conversion, table bounds,
cycle detection, and marker validation. No unsafe access, unchecked index, or
source/provider freshness change is justified by a cold-layout hypothesis.

## Existing tests and required guards

The existing tests are sufficient for the intended collector-only helper
move. `scratch_differential_matches_owned_chain_helper_and_resets` compares
the reusable collector with the owned reference for valid, empty,
invalid-start, invalid-index, cyclic, early-end, late-end, invalid-marker,
short-table, and non-ENDOFCHAIN empty cases. It compares exact error strings,
checks successful sector order and logical visited length, verifies every
error leaves an empty result and zero logical visited length, and repeats
successful cases while checking retained pointers and capacities.

The surrounding regression tests add the relevant state and allocation
boundaries: `reusable_chain_scratch_reuses_buffers_for_different_lengths`,
`reusable_chain_scratch_reserves_growth_before_walking`,
`reusable_chain_scratch_resets_after_success_and_empty_chain`,
`reusable_chain_scratch_resets_after_cycle_and_reuses_after_error`,
`reusable_chain_scratch_preserves_early_and_late_end_errors`,
`rejects_chain_lengths_before_reserving_chain_storage`, and
`malformed_large_declarations_do_not_unwind`. Together they guard reset,
reuse, preflight order, terminal-marker precedence, and panic resistance
without depending on generated code layout.

There is no direct `CheckedBitSet` unit test. That omission is acceptable if
the final patch only moves `collect_exact` diagnostic formatting and does not
change bitset representation or method bodies: the differential suite covers
all bitset operations reached by collection, and the allocator lane later
guards successful-path allocation counts and bytes. If the candidate changes
`CheckedBitSet::insert` or `contains`, add one private table-driven boundary
guard before candidate measurement. It should cover logical lengths 0, 1, 64,
and 65; valid bits 0, 63, and 64 where applicable; out-of-range `contains`
returning false; successful and duplicate insertion; the exact out-of-range
error; and unchanged words after an invalid insertion. A test should exercise
the missing-backing-word diagnostic only if the patch changes that invariant
or makes it newly reachable. Do not add timing assertions or allocator
failure simulation to this unit test.

No new differential test is needed merely to assert `#[cold]` or
`#[inline(never)]`; assembly and symbol inspection on the measured binary
must establish those code-layout facts. If the final patch introduces a
different helper boundary or changes the collector's control flow, the tests
agent must first retain the existing table-driven differential coverage and
add the smallest case for every newly reachable branch before capture.

## ADR and measurement compatibility

The design remains compatible with accepted constraints when it stays private
and layout-only:

- ADR 0001, 0002, and 0024 require the existing private `litchi-cfb` owner and
  current OLE2/OOXML topology; no public API, dependency, or crate boundary
  may change.
- ADR 0003 and 0006 require deterministic typed errors, validation order,
  immutable failure behavior, and exact preservation semantics. The listed
  messages and reset state are therefore behavioral obligations, not optional
  diagnostics.
- ADR 0005 requires profile and statistical evidence, explicit resource
  accounting, and fallible bounded allocations. A cold attribute is a
  hypothesis; it is not a latency claim or permission to remove validation.
- ADR 0008 requires source, binary, plan, profile, assembly, and test custody
  to bind to the measured candidate. The same binary must provide the
  `collect_exact`/`CheckedBitSet` symbols used for instruction review.
- ADR 0010, 0011, and 0026 preserve package, physical CFB, and directory
  ownership boundaries. Collect-before-claim and physical reconciliation
  remain required positive work.

The frozen 0536 plan's native admission rule requires every primary XLS p50
to improve by at least 3% in both paired repeats, while preserving its CFB
shape guards and correctness/quality checks. Allocation calls, bytes, and
incremental peak must not grow materially. Review matched adverse rows and
same-build variation over 5%, retain all raw rows, and keep setup CFB dumps
separate from timed constructor dumps. Assembly must demonstrate that the
ordinary loop has no new call or allocation and that only rare diagnostic
tails moved. A lower profile share or shorter helper is mechanism evidence,
not an adoption result.

The 0524 measured fused visited-bit implementation remains rejected as a
specific failed experiment; that does not categorically reject every future
visited representation. The 0279 change concerned provider per-read versus
operation-scoped source freshness, not visited-bit state; its performance
drift gate failed while correctness passed. Neither direction is in scope for
this cold-error-helper probe.

## Disposition

The retained source satisfies the reviewed error, order, allocation, reset,
freshness, and ownership contracts. Existing collector differential and
scratch-reuse tests are the minimum sufficient guard for a formatting-only
candidate. A direct bitset boundary test becomes a pre-measurement
requirement only if the final patch changes bitset methods or representation.
The inspected candidate diff satisfies the source-bound conditions and may be
staged for the frozen measurement protocol; no production optimization or
performance claim is approved by this review. OLE2/OOXML remains the active
priority, ODF is deferred, and iWork is outside scope.
