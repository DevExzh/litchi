# 0535 CFB collector and checked-bitset source review

`scope: read-only OLE2/CFB source and contract review before any 0535 implementation`

`revision: bb5bfaa7be4fdffeae3fddaee3ed1266c3417520`

`reviewed source: crates/litchi-cfb/src/file.rs SHA-256
72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b`

`performance_claim: none`

This review covers [`CheckedBitSet`](../../../../crates/litchi-cfb/src/file.rs#L22),
[`SectorChainScratch`](../../../../crates/litchi-cfb/src/file.rs#L2764),
`SectorChainScratch::collect_exact`, its stream-allocation callers, and the
existing owned-helper differential tests. No Rust source, test, build,
benchmark, profiler, or capture was run or changed for this review. The
current production collector is baseline behavior after the earlier rejected
visited-bit experiment.

## Source boundary

`CheckedBitSet` is a private, fallible equivalent of the checked portion of a
fixed bit set. It stores zero-initialized `u64` words and a logical `bit_len`.
`try_with_capacity` computes the ceiling word count, performs
`try_reserve_exact`, maps the failure through the supplied allocation
resource label, and then zero-initializes the words. `contains` is already
marked `#[inline]`; it returns `false` for a bit outside `bit_len` and uses a
checked `get` for the word. `insert` checks the same boundary, uses checked
`get_mut`, sets the bit, and returns `Ok(())` on success. The set operation is
idempotent for a repeated bit.

`SectorChainScratch` owns one retained sector vector and one retained visited
map. `validate_stream_allocations` creates separate scratch values for the
MiniFAT and FAT namespaces because their table lengths can differ; a shared
map would retain the larger namespace in the smaller path. The general
`collect_sector_chain_exact` helper remains an owned-result reference
implementation. `collect_exact` is a private reuse path and must not be
treated as permission to change chain ownership, physical-role claiming, or
the two sector namespaces.

## Exact bitset obligations

The following behavior is observable through chain errors and must survive
any representation or layout change:

| Operation | Required behavior |
| --- | --- |
| `try_with_capacity(bit_len, resource)` | Allocate `bit_len.div_ceil(64)` words fallibly with `resource`, then zero-initialize them. |
| `contains(bit)` with `bit >= bit_len` | Return `false` without an error or allocation. |
| `contains(bit)` within `bit_len` | Read the checked word and mask; stale bits must not become visible after a scratch reuse. |
| `insert(bit)` with `bit >= bit_len` | Return `OleError::CorruptedFile("bit index {bit} exceeds checked bit-set capacity {bit_len}")`. |
| `insert(bit)` with no backing word | Return `OleError::CorruptedFile("bit index {bit} has no backing word")`; this is an internal-invariant guard. |
| successful `insert` | Set only the requested bit, without allocating. |

`SectorChainScratch::prepare_visited` can retain more words than the current
table needs. It sets the new logical length and fills all retained words with
zero on every nonempty collection. `reset` only clears the sector vector and
sets `bit_len` to zero; it intentionally retains capacities and word storage.
Consequently, replacing the full fill with a partial clear would require a
proof that a later larger table can never expose an uncleared tail. It is not
a first candidate without that proof and an allocation/reuse guard.

## `collect_exact` behavior and error order

The collector has one reset at entry and a second reset whenever the closure
returns an error. Its required order is:

1. Clear the prior result and make the visited map logically empty.
2. For an empty declaration, accept only `start_sector == ENDOFCHAIN`.
3. For a nonempty declaration, reject a start marker at or above
   `MAXREGSECT`, then reject `expected_count > allocation_table.len()`.
4. Reserve the sector vector first with the `sector-chain entries` resource,
   then grow and clear the visited map with the `sector-chain map` resource.
5. For each of exactly `expected_count` entries, safely convert the sector to
   `usize`, check the table bound, check `contains`, call `insert`, push the
   sector in chain order, and read the next marker through `get`.
6. On the final entry require `ENDOFCHAIN`. On an intermediate entry reject
   early `ENDOFCHAIN`, reject a marker at or above `MAXREGSECT`, and continue
   with the next sector.
7. On every error, clear the result and set the logical visited length to zero;
   retain allocated capacities. On success, expose exactly the declared
   sectors and leave `visited.bit_len == allocation_table.len()` for a
   nonempty chain (zero for an empty chain).

The exact formatted diagnostics are part of the compatibility contract:

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

The invalid-index text is used both by the checked conversion/bounds paths
and by the defensive `allocation_table.get` path. The intermediate marker
checks occur only after the current sector has been checked and recorded, so
changing the order can change which malformed input wins. `Vec::push` is safe
without another fallible operation because the exact sector reservation is
performed before the loop; a candidate must not remove or weaken that
reservation proof.

## Allocation, reset, and namespace contracts

On a successful call, all possible growth occurs in the two fallible
reservations before the loop. A later call may reuse both buffers without
allocation, including after an error. A sector reservation failure uses the
`sector-chain entries` label; map growth failure uses `sector-chain map`.
Those labels and the reservation order are part of the malformed-input and
allocation-failure behavior. A candidate must not replace either with an
infallible reserve, an intermediate collection, or a fresh per-stream map.

The scratch result is private but its reset state affects subsequent stream
validation. A failed MiniFAT collection must leave no sectors that a later
claim could observe, and a failed regular collection must not inherit visited
bits from a MiniFAT or an earlier regular stream. The separate scratch values
and the `fill(0)` in `prepare_visited` provide that namespace and freshness
boundary. `claim_chain` still runs only after a complete collection succeeds;
the collector change must not publish physical ownership while walking.

## Existing differential and regression coverage

The current tests already exercise the main contract without depending on a
particular optimized representation:

| Test | Coverage that must remain valid |
| --- | --- |
| `scratch_differential_matches_owned_chain_helper_and_resets` | Two passes compare scratch output or exact error text with `collect_sector_chain_exact` for valid, empty, cyclic, invalid-start, invalid-index, early-end, late-end, invalid-marker, short-table, and non-ENDOFCHAIN empty cases. Every error checks an empty result and zero logical visited length; successful cases check output and table-length state; repeated valid cases check retained pointers and capacities. |
| `reusable_chain_scratch_reuses_buffers_for_different_lengths` | Different FAT/MiniFAT-sized tables retain sector and word storage while updating the logical bit length. |
| `reusable_chain_scratch_reserves_growth_before_walking` | A larger requested result reserves before walking; an early-end failure leaves the result empty and retains the larger sector capacity. |
| `reusable_chain_scratch_resets_after_success_and_empty_chain` | Explicit reset and a valid empty chain both leave no sectors and zero logical visited length. |
| `reusable_chain_scratch_resets_after_cycle_and_reuses_after_error` | A cycle reports the exact error, resets state, and permits a subsequent valid collection. |
| `reusable_chain_scratch_preserves_early_and_late_end_errors` | Early and late terminal-marker errors retain their current precedence and exact wording. |
| `rejects_chain_lengths_before_reserving_chain_storage` | The owned helper rejects an overlong declaration before reserving result storage, preserving the reference allocation boundary. |
| `malformed_large_declarations_do_not_unwind` | Large malformed declarations remain error paths rather than panics. |

There is no direct unit test for `CheckedBitSet`'s boundary methods. That is
acceptable for a collector-only code-layout change: the differential test
exercises the collector's bitset operations while comparing against the owned
chain helper, although it is not an independent oracle for internals shared
by both helpers. If a future candidate changes the bitset implementation or
representation, the smallest useful additional guard is one table-driven
private test covering
logical lengths 0, 1, 64, and 65; valid and out-of-range `contains`; bits 0,
63, and 64 where valid; duplicate insertion; the exact out-of-range `insert`
error; and unchanged state after an invalid insertion. Allocation timing or
allocator-count assertions do not belong in that unit test; the existing
allocation lane is the guard for those properties.

## Candidate boundary and rejected directions

The current source has a narrow code-layout opportunity if a profile shows
that bitset error formatting is affecting the ordinary path: retain the two
separate `contains` and `insert` calls, move the two rare formatted `insert`
errors into private `#[cold] #[inline(never)]` helpers, and give only the
successful insert body an ordinary inline opportunity. The helpers must take
copied values, preserve the exact strings above, and leave the bounds checks,
error order, reset behavior, and fallible reservations unchanged. `contains`
already has an ordinary inline hint. The retained profiles do not identify a
separate positive hot `insert` callee, so cold splitting remains a hypothesis
pending a fresh instruction map, matched measurement, and generated-code
review; it is not a speedup claim.

The 0524 experiment measured one private implementation that fused the
checked visited-bit lookup and set, removed duplicate checked bit access, and
failed the native admission rule. That result rejects retaining that measured
implementation; it does not categorically reject every different visited-bit
representation or every future duplicate-access reduction. Those designs are
outside this diagnostic scope. A genuinely distinct proposal would need its
own source proof for the exact contracts above, independent differential
coverage, and fresh native/allocation evidence. The retained differential test
does not by itself reopen the 0524 result.

The 0279 proposal was a provider per-read versus operation-scoped source
freshness session, not a visited-bit cache or cross-call collector state. It
was reverted because its strict ABBA performance drift gate failed; its
correctness review passed, and the repository retains per-read freshness. No
provider or freshness change is in scope here. Likewise, no candidate may
change `claim_chain`, physical-role ownership, FAT/MiniFAT namespace
separation, stream validation order, or physical reconciliation. Those are
required work, not removable collector overhead.

## Performance guard implications

The collector's profile share identifies executed work; it does not establish
that the work can be removed. Any candidate staged after this review must use
the frozen native admission protocol: all four primary XLS workflows must
clear the required p50 improvement threshold in both paired repeats, while
the tiny, many-small, few-large CFB shapes remain individual guards. Native
constructor timing is the primary signal; allocator elapsed time is a
separate lane.

The allocation lane must retain identical allocation-call, allocated-byte,
and incremental-region-peak behavior on successful paths, with no new
per-stream allocation. Matched and same-build latency, mean/tail, RSS, and
outlier records remain reviewable; a large profile or assembly improvement
cannot waive an adverse row. Assembly and source manifests must bind the
measured binary to the exact candidate source, and any candidate that is
rejected must restore production runtime behavior while retaining only
representation-independent tests.

## Source disposition

The current collector and bitset satisfy the reviewed correctness,
allocation, reset, and freshness contracts. There is no correctness blocker
to staging a narrowly measured cold-error/layout experiment, but there is no
0535 production change or performance claim to approve here. Preserve the
existing differential tests as the minimum regression guard; add the direct
bitset boundary test only if the bitset itself changes. OLE2/OOXML remains the
active optimization priority, with ODF deferred and iWork outside scope.
