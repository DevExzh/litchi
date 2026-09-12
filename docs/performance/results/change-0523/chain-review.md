# 0523 CFB chain-owner review

`scope: read-only CFB/OLE2 hotspot review`

`performance_claim: none`

Reviewed against the production `litchi-cfb` owner at the current worktree.
This review proposes one bounded experiment for the next CFB measurement lane;
it does not edit production code, add tests, or report a new measurement. The
fresh 0523 operation profile and allocator capture must run before a candidate
is applied. OLE2/OOXML remains ahead of ODF in the active optimization queue.

## Governing constraints

The accepted ADR set and index were read before selecting this candidate. The
direct constraints are [ADR 0001](../../../adr/0001-priorities-and-api-layers.md),
[ADR 0002](../../../adr/0002-crate-topology.md),
[ADR 0003](../../../adr/0003-snapshots-edits-and-patches.md),
[ADR 0005](../../../adr/0005-io-memory-and-performance.md),
[ADR 0006](../../../adr/0006-validation-security-and-compatibility.md),
[ADR 0008](../../../adr/0008-migration-and-verification.md),
[ADR 0010](../../../adr/0010-facade-archive-ownership.md),
[ADR 0011](../../../adr/0011-ooxml-physical-package-ownership.md),
[ADR 0024](../../../adr/0024-current-topology.md), and
[ADR 0026](../../../adr/0026-ole-directory-metadata-binding.md). They require
the CFB owner to retain fallible bounded allocations, typed refusals, source
and physical validation, lossless behavior, and its current crate boundary.
The performance record in [ADR 0005](../../../adr/0005-io-memory-and-performance.md)
also requires a representative matched measurement before retaining a speed
or memory claim.

## What the retained 0511 profiles actually show

The 0511 source-open **control** in
[`before/callgrind-exclusive.txt`](../change-0511/before/callgrind-exclusive.txt)
collected 15,127,968 simulated instruction references over five
`SourceBackedWorkbook::from_read_at_with_limits` calls. Its exclusive shares
were:

| Owner | Control Ir | Share |
| --- | ---: | ---: |
| `SectorChainScratch::collect_exact` | 5,601,140 | 37.03% |
| `OleFile::claim_sector` | 2,489,550 | 16.46% |
| `validate_physical_sector_layout` | 1,991,710 | 13.17% |
| `validate_stream_allocations` | 1,979,410 | 13.08% |
| `load_fat` | 1,407,020 | 9.30% |

The 0511 source-open **final** profile in
[`after/callgrind-exclusive.txt`](../change-0511/after/callgrind-exclusive.txt)
is the post-FAT-batching binary, not a current 0523 baseline. It collected
13,973,601 references with the unchanged chain and ownership work:

| Owner | Final Ir | Share |
| --- | ---: | ---: |
| `SectorChainScratch::collect_exact` | 5,601,140 | 40.08% |
| `OleFile::claim_sector` | 2,489,550 | 17.82% |
| `validate_physical_sector_layout` | 1,991,710 | 14.25% |
| `validate_stream_allocations` | 1,979,410 | 14.17% |
| `load_fat` | 250,820 | 1.79% |

The old 9.30% `load_fat` figure therefore cannot select another FAT
optimization. The retained final binary is useful for identifying the
residual owners, but compiler, host, and worktree changes require a fresh
0523 profile before ranking is treated as current. The separate final
few-large CFB guard profile also leaves `collect_exact` largest (5,571,945
references, 42.60% of that guard runner), but that runner includes its timers,
oracles, drops, and result construction and is not an operation-local source
profile.

## Selected candidate: fuse the visited-bit lookup and set

The current reusable collector at
[`file.rs#L2778`](../../../../crates/litchi-cfb/src/file.rs#L2778) validates a
sector index and then performs two checked operations on the same visited bit:

```text
if self.visited.contains(slot) { cycle error }
self.visited.insert(slot)?
```

The next bounded experiment should add one private
`CheckedBitSet::test_and_set(bit) -> Result<bool, OleError>` operation and use
it only in `SectorChainScratch::collect_exact`:

```text
if self.visited.test_and_set(slot)? { cycle error }
```

The operation should retain the existing bit-length check and backing-word
lookup, return whether the bit was already set, and OR the bit into the same
word. It should be marked for inlining if the fresh compiler profile shows the
private call boundary is material. `contains` and `insert` remain available to
their other owners; this experiment must not broaden their semantics or touch
the general `collect_sector_chain_exact` helper.

This removes the repeated word calculation, mask calculation, and checked
lookup in the common collector loop. It does not remove cycle detection,
allocation, bounds, marker, ownership, or physical-layout validation. It adds
no state, allocation, dependency, public API, unsafe code, or source-I/O
behavior. Among the larger residual loops this is the strongest safe first
candidate because it removes provably duplicate work without attempting to
fuse validation phases or weaken a hostile-input boundary.

## Exact safety and error-order proof

`collect_exact` first calls `reset`, handles the empty-chain case, rejects an
invalid start marker, and rejects `expected_count > allocation_table.len()`.
It then reserves `sectors` fallibly under `"sector-chain entries"`, prepares
the visited map fallibly under `"sector-chain map"`, and sets
`visited.bit_len = allocation_table.len()` before the walk. Those operations
and their order stay unchanged.

For every loop iteration, the collector converts the current `u32` sector to
`usize` and checks `slot < allocation_table.len()` before touching the visited
map. Since `prepare_visited` sets the bit length to that same table length, the
new method's out-of-range branch is unreachable on a valid collector state,
but it must retain the current typed error:

```text
bit index {bit} exceeds checked bit-set capacity {bit_len}
```

The method must also retain the current missing-word error:

```text
bit index {bit} has no backing word
```

For a previously clear bit, `test_and_set` produces exactly the word state
that `contains` followed by `insert` would produce. For a repeated bit, the OR
operation leaves the already-set word unchanged; the collector returns the
same cycle error before appending a sector or reading its next marker. On any
bit-set error, the method must not mutate the word. `collect_exact` still
clears its scratch state after every error, so the post-error state and buffer
reuse contract remain unchanged.

The surrounding order remains:

1. invalid sector conversion and allocation-table bounds;
2. cycle detection;
3. the already-established exact `Vec` push into fallibly reserved capacity;
4. allocation-table lookup;
5. early/late `ENDOFCHAIN` and invalid marker checks; and
6. reset of scratch state after any failure.

The candidate therefore preserves every reachable `OleError` message and
precedence in the collector. It does not change the resource labels or move
either fallible reservation. In particular, it does not turn a malformed
chain into an infallible walk and does not make `Vec::push` depend on a new
capacity assumption.

## Why the other residual loops stay out of this experiment

`validate_stream_allocations` calls `collect_exact` and then
[`claim_chain`](../../../../crates/litchi-cfb/src/file.rs#L1051) for regular
streams. Fusing role writes into the chain walk would change error precedence:
the current collector reports a late cycle, short chain, or marker error
before it reports an overlap from `claim_sector`. A fused walk could report
the overlap first and could leave partial role mutations on a failed walk.
The second ownership pass remains required until a separate representation and
rollback/error-order proof exists. The `claim_sector` successful path also
must retain checked conversion, physical-file bounds, role labels, and exact
overlap diagnostics; its 0511 share is not evidence that any of those checks
are redundant.

`validate_physical_sector_layout` at
[`file.rs#L1157`](../../../../crates/litchi-cfb/src/file.rs#L1157) cannot be
deleted or skipped because unclaimed physical sectors are valid when their FAT
entry is `FREESECT`. The final loop is what detects an unclaimed non-free
marker and reports its physical index in ascending order. Moving that check
into `load_fat` would require retaining a bounded bad-sector index and
carefully preserving later claims plus the first-error ordering; a counter or
an early generic error would not preserve the current contract. The accepted
0511 record explicitly leaves this reconciliation mandatory.

## Existing correctness guards

The candidate must be compared against the unchanged helper behavior and run
through the existing CFB owner/consumer coverage, including:

- `file::tests::reusable_chain_scratch_reuses_buffers_for_different_lengths`;
  `reusable_chain_scratch_reserves_growth_before_walking`;
  `reusable_chain_scratch_resets_after_success_and_empty_chain`;
  `reusable_chain_scratch_resets_after_cycle_and_reuses_after_error`; and
  `reusable_chain_scratch_preserves_early_and_late_end_errors`;
- `file::tests::rejects_chain_lengths_before_reserving_chain_storage`,
  `detects_cycles_in_allocation_chains`, malformed count/no-unwind cases, and
  the direct FAT/MiniFAT range and truncation cases;
- `allocation_validation_tests::open_rejects_minifat_cycles_and_invalid_markers`,
  `open_rejects_fat_stream_cycles_and_invalid_markers`, the short/excess
  terminal-marker tests, hostile count tests, overlap tests, and marker tests;
- `validation_tests::fat_topology_rejection_is_a_deterministic_structured_issue`;
  and
- the writer reopen boundaries for mixed MiniFAT/FAT storage, DIFAT boundary
  streams, and the exact FAT boundary.

A candidate-only differential test should feed both
`collect_sector_chain_exact` and the scratch collector valid chains, empty
chains, cycles, invalid indexes, early/late terminal markers, invalid markers,
and overlong declarations. It should compare the returned sectors or exact
`OleError` debug/display text and assert scratch cleanup after failures. This
is a semantic differential guard, not a new claim about malformed-input
coverage; no such test is added in this read-only review.

## Measurement gate

Before applying the candidate, capture a fresh current-source CFB profile with
separate `collect_exact`, `claim_sector`, `validate_stream_allocations`, and
`validate_physical_sector_layout` attribution. Pair it with operation-local
allocator counts/bytes and incremental live-peak metrics over the existing
FAT-heavy XLS source-open case, plus tiny MiniFAT and few-large regular-FAT
guards. The allocator and Callgrind lanes must retain their distinct scopes.

If applied, use serial ABBA release captures with the same input, source
read, output, and error oracles. Retain the change only if total end-to-end
results improve repeatably within the declared drift/adverse thresholds and
all CFB/XLS/DOC/PPT correctness and strict quality gates pass. A lower
`collect_exact` instruction count alone is diagnostic; it does not establish a
latency, allocation, RSS, physical-I/O, cold-cache, provider, scaling, or
native Office claim. If the fresh profile shows the fused operation is below
noise or the candidate regresses the total workload, close it as a rejected
experiment and retain the evidence without changing production.

No ODF work is selected by this review.
