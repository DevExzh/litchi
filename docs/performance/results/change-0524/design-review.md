# 0524 CFB visited-bit candidate design review

The candidate is recorded in [`candidate.patch`](candidate.patch). It is an
artifact only and has not been applied to the production checkout. The patch
targets the current `crates/litchi-cfb/src/file.rs` blob at revision
`477281a2f83c3256bf3cc06fbc5c57724c9b6bbc` (blob SHA-1
`1da9ad4f2c0b6989e3d5e7828cb7fc964b5702ec`). No production source, test, build,
profile, or capture was changed or run for this review.

## Applicable decisions and scope

The accepted ADR index and the accepted records were read before preparing the
artifact. The complete indexed set is bound by
[`change-0523/adr-manifest.json`](../change-0523/adr-manifest.json), whose
revision is `0e6379307c5174babfc832edc6113a16c8e9233a`; this review applies the
constraints from ADR 0001 (strict typed layers and safety), ADR 0002 (the
`litchi-cfb` crate boundary), ADR 0003 (fallible, non-mutating state contracts),
ADR 0005 (bounded allocations and measured performance), ADR 0006 (validation,
preservation, and hostile-input behavior), ADR 0008 (verification gates), ADR
0010/0011 (container ownership), ADR 0024 (current topology), and ADR 0026
(inert shared OLE metadata). No accepted ADR is amended by this candidate.

This is one private implementation experiment inside the existing CFB owner.
It adds no public API, dependency, allocation, storage, source-I/O, unsafe-code,
or validation-phase change. OLE2/OOXML remains the active performance priority;
ODF work stays deferred until that optimization goal is complete.

## Measured owner and motivation

The fresh 0523 profile is the selection evidence. Its operation owner for the
XLS source-open profile is
`litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at_with_limits`;
the positive incoming caller is the existing
`SourceBackedWorkbook::from_read_at` wrapper. The CFB owner is
`litchi_cfb::file::OleFile<R>::open`, and its timed profile excludes fixture
generation, the file-size oracle, report construction, and object drop. The
profile analysis passed its source/binary binding, incoming-caller,
single-constructor, final-dump, and identity checks.

Across the two XLS profile repeats, the disjoint exclusive shares are:

| Owner | Exclusive share of the constructor |
| --- | ---: |
| `SectorChainScratch::collect_exact` | 40.08–40.09% |
| `OleFile::claim_sector` | 17.81–17.82% |
| `validate_physical_sector_layout` | 14.25–14.26% |
| `validate_stream_allocations` | 14.16–14.17% |
| `load_fat` | 1.79–1.80% |

These are Callgrind Ir shares, so they identify work and ownership but do not
predict elapsed time, allocation counts, memory use, or end-to-end benefit.
The retained raw evidence is [`profile-analysis.json`](../change-0523/profile-analysis.json)
and the per-profile self reports under
[`change-0523/baseline`](../change-0523/baseline/).

The direct CFB guards show the same owner at different scales. In the retained
timed dump, `collect_exact` accounts for 1,114,389 of 2,614,530 Ir (42.62%) in
the few-large shape, 152,897 of 2,805,468 Ir (5.45%) in many-small, and 1,040
of 48,868 Ir (2.13%) in tiny. The many-small dump records 81,920 Ir in
`CheckedBitSet::insert` from the four thousand ninety-six sector-chain map
updates under `validate_stream_allocations`; other `insert` work remains in
directory and ownership paths. This makes the duplicate lookup/set in the
reusable collector a bounded candidate while retaining tiny and many-small as
guards. The 0523 evidence is attribution only; it is not a before/after claim.

## Proposed operation

The patch adds this private operation to the existing `CheckedBitSet`:

```rust
#[inline]
fn test_and_set(&mut self, bit: usize) -> Result<bool, OleError>
```

It performs the same checked bit-length and backing-word checks as `insert`,
reads the previous mask state, ORs the mask into the word, and returns whether
the bit was already set. `#[inline]` is intentional: the fresh profile shows a
material private `insert` boundary in the selected collector, and this
operation should let the ordinary release compiler combine the one lookup and
one mask calculation with the loop. The attribute is an optimization hint,
not a measured performance result.

Only `SectorChainScratch::collect_exact` uses the operation:

```rust
if self.visited.test_and_set(slot)? {
    return Err(OleError::CorruptedFile(format!(
        "Cycle detected in {table_name} chain at sector {sector}"
    )));
}
```

The general `collect_sector_chain_exact` helper continues to use
`contains` followed by `insert`. Directory traversal, directory ownership,
MiniFAT ownership, and every other `CheckedBitSet` user are unchanged. The
candidate therefore removes only the repeated word/mask computation in the
reusable exact collector.

## Correctness and error-order proof

`collect_exact` still begins by resetting scratch state, handles the empty
chain, rejects an invalid start marker, rejects a declared count larger than
the allocation table, reserves `sectors` under `"sector-chain entries"`, and
prepares the visited map under `"sector-chain map"`. The candidate does not
move either fallible reservation or alter `prepare_visited`, `reset`, or buffer
reuse.

For each loop iteration, the existing checked `u32` to `usize` conversion and
allocation-table bound check run before the visited operation. On the valid
collector path, `prepare_visited` gives the bit set exactly the same bit length
as the allocation table. The candidate nevertheless retains both existing
typed diagnostics in `test_and_set`:

```text
bit index {bit} exceeds checked bit-set capacity {bit_len}
bit index {bit} has no backing word
```

The first is unreachable after the surrounding slot check but remains part of
the private method's checked contract. The second is checked before the word is
mutated. Both errors remain `OleError::CorruptedFile` with the same text as
`insert`.

For a clear bit, the new operation leaves exactly the word state produced by
the old `contains` then `insert` sequence. For a set bit, OR leaves the word
unchanged and the collector immediately returns the same cycle diagnostic,
before appending the sector or reading its next marker. A missing backing word
or out-of-range bit returns before mutation, just as the old sequence did after
`contains` returned false. Any later marker or table error still follows the
same set-bit mutation, and the existing error path clears all scratch state
through `reset`.

The remaining order is therefore unchanged:

1. checked sector conversion and allocation-table bounds;
2. cycle detection;
3. the already reserved exact `Vec` push;
4. allocation-table lookup;
5. early/late `ENDOFCHAIN` and invalid-marker checks; and
6. scratch reset after any failure.

The operation performs no allocation and cannot turn a fallible boundary into
an infallible one. `contains` and `insert` keep their existing semantics for
all other owners. No public type, archive grammar, physical-sector check,
stream-ownership check, or source provider is exposed or bypassed.

## Required validation before adoption

The artifact needs a candidate-only differential guard before production
retention. It should exercise the general exact helper and scratch collector on
valid chains, empty chains, cycles, invalid sector indexes, early and late
terminal markers, invalid markers, and overlong declarations; compare sector
results or exact `OleError` display/debug text; and assert scratch cleanup after
each failure. No test is added in this design-only artifact.

The existing CFB owner and consumer suites must remain green, including
reusable scratch growth/reset/cycle/marker tests, hostile count and malformed
FAT/MiniFAT tests, overlap and topology validation, direct range/truncation
checks, and writer reopen boundaries for MiniFAT/FAT, DIFAT, and exact FAT
limits. Strict formatting, Clippy, rustdoc, workspace boundary, and claim
registry gates remain required. Native producer, fuzz, physical-provider,
cold-cache, concurrency-scaling, and broad CRUD requirements are not inferred
from this private candidate.

## Measurement gate

Apply the patch only to an isolated candidate source copy after the 0523
baseline is retained. Run the same release matrix, corpus manifests, source
and output oracles, allocator envelope, and profile boundaries. Use serial
matched ABBA repetitions for the CFB tiny, many-small, and few-large guards and
the XLS source-open owner. Compare end-to-end p50/p95/p99 and recomputed
intervals, allocation calls/bytes, live-byte balance and incremental peak, RSS,
and any available whole-child counters. Retain all observations, including
same-build changes above the declared drift threshold.

A lower `collect_exact` Ir count is diagnostic only. Retain the candidate only
if the named end-to-end OLE2/OOXML workflows improve repeatably within the
declared adverse thresholds, with no correctness, preservation, typed-error,
allocation, peak-memory, RSS, or source-boundary regression and with all
quality gates passing. Otherwise record the rejected experiment and keep the
current production implementation. No speedup, allocation reduction, or
10x claim is made by this review.
