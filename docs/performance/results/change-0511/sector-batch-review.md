# 0511 sector-batched FAT decode review

Reviewed 2026-09-11 against base revision
`b0b9ebb577d616a9a0ed93f16619a1b79f3b3bc7`. The owner has applied the
sector-batched FAT decoder in [file.rs](../../../../crates/litchi-cfb/src/file.rs#L841-L876);
this review makes no production edit. The current file.rs SHA-256 is
`5f70ab560dd2d954e487fabeaac19d01d6b6158d74a6576834e01940d2cf4e72`, matching
the current [0511 source manifest](source-manifest.json).

This is the pre-measurement source review. Final measured evidence and gate
disposition are recorded in [0511](../../changes/0511-cfb-fat-entry-reservation.md)
and [final-review.md](final-review.md). The prospective wording below records
the review state when the candidate was selected.

## Disposition

The candidate is mechanically sound and bounded for a controlled test and
measurement. It is not retained as a performance result yet: the final normal
build is still running, and no native elapsed-time, allocation, or RSS capture
has been recorded for this candidate.

The preceding push-only variant remains a rejected diagnostic. Its paired
Callgrind artifacts are preserved under [initial-push/](initial-push/): total
instruction references rose from 15,127,968 to 15,462,114 (+2.21%), while
exclusive load_fat references rose from 1,407,020 to 1,740,600 (+23.71%).
Those are simulated instruction counts, not a general end-to-end latency
regression. The rejection is recorded in
[initial-push/rejection.json](initial-push/rejection.json).

The next candidate changes only the FAT payload decode loop, leaving header and
DIFAT collection, MiniFAT construction, directory parsing, stream validation,
and physical ownership checks out of scope:

```rust
fat.extend(
    sector_data
        .as_chunks::<4>()
        .0
        .iter()
        .copied()
        .map(u32::from_le_bytes),
);
```

## Capacity proof

The existing exact reservation and count checks are sufficient for this
infallible extend:

1. CFB header validation admits only 512- or 4096-byte sectors. Both sizes
   are divisible by four, and sector_data is a zero-filled stack buffer
   sliced to exactly self.sector_size.
2. Header/DIFAT traversal completes before the payload loop and checks
   fat_sectors.len() == expected_fat_sectors.
3. The code computes
   fat_entry_count = fat_sectors.len() * (self.sector_size / 4) with checked
   multiplication, then successfully calls
   try_vec_with_capacity(fat_entry_count, "FAT entries").
   The successful reservation gives fat capacity at least the total count.
4. as_chunks::<4>().0 contains exactly self.sector_size / 4 complete
   [u8; 4] arrays per sector. iter(), copied(), and map preserve that
   cardinality and expose an ExactSizeIterator with equal lower and upper
   bounds.
5. After k outer iterations, fat.len() is exactly
   k * fat_entries_per_sector. The next extend adds one sector's exact
   number of entries, so its end length never exceeds fat_entry_count.

Therefore Vec::extend has enough existing capacity for every sector and does
not call the allocator. The fallible "FAT entries" reservation remains the
only allocation boundary. This proof must remain adjacent to the loop: if a
future change permits partial chunks, adds another push, changes sector
geometry, or weakens the declared-sector count check, the infallible extend
must be withdrawn or replaced with a fallible growth path.

## Decode and error-order proof

The candidate preserves all reachable parser behavior under those invariants:

- read_sector_into still runs once per FAT sector, in the same order, before
  decoding that sector. Its existing zero-fill behavior for a truncated final
  sector and its typed error for a sector beginning outside the physical file
  remain unchanged.
- Every mapped item is a complete [u8; 4]. The current
  [read_u32_le](../../../../crates/litchi-cfb/src/file.rs#L2541-L2546)
  helper's try_into failure branch is unreachable for this fixed-chunk source.
  u32::from_le_bytes performs the same little-endian conversion on the same
  four bytes, so removing that unreachable InvalidFormat branch does not alter
  a reachable malformed-input result.
- The FATSECT and DIFSECT ownership-marker checks remain after all entries are
  decoded, and self.fat is still published only after those checks succeed.
  The decoded table order and values are unchanged.
- The exact reservation remains fallible and keeps its "FAT entries" label and
  position before the first FAT-sector read. Since extend cannot grow under the
  capacity proof, it introduces no infallible allocation on the current path.

This reasoning applies only to the fixed as_chunks::<4>() loop. It does not
authorize changing the variable-length FAT/DIFAT location vectors, the generic
exact chain helper, directory queues, or the MiniFAT loop. MiniFAT still lacks
an isolated profile contribution and remains deferred.

## Required validation before retention

The existing focused malformed, ownership, chain, and truncation tests remain
the regression surface:

- file::tests::fat_stream_*,
  file::tests::minifat_stream_*, file::tests::direct_ranges_*,
  file::tests::detects_cycles_in_allocation_chains,
  file::tests::rejects_self_referential_difat_chains,
  file::tests::rejects_difat_counts_beyond_the_physical_file_before_reserving,
  file::tests::read_sector_into_zero_fills_a_truncated_final_sector, and
  file::tests::tolerates_nonfree_fat_padding_beyond_the_physical_file;
- allocation_validation_tests::hostile_count_metadata_is_rejected_before_index_allocation,
  allocation_validation_tests::hostile_difat_and_minifat_counts_are_rejected_before_index_allocation,
  allocation_validation_tests::rejects_directory_and_fat_sector_overlap,
  allocation_validation_tests::rejects_fat_sector_without_fatsect_marker,
  allocation_validation_tests::open_rejects_fat_stream_cycles_and_invalid_markers,
  allocation_validation_tests::open_rejects_fat_stream_short_excess_and_terminal_markers,
  and allocation_validation_tests::still_rejects_a_header_difat_list_shorter_than_its_count;
- validation_tests::fat_topology_rejection_is_a_deterministic_structured_issue;
- writer reopen boundaries:
  sequential_writer::mixed_minifat_fat_nested_storage_clsids_reopen,
  sequential_writer::difat_boundary_streams_without_retaining_source_payload,
  and sequential_writer::exact_fat_boundary_has_no_difat_and_one_over_has_difat.

The exact scoped commands are:

```text
cargo test --locked -p litchi-cfb --lib 'file::tests::fat_stream_'
cargo test --locked -p litchi-cfb --lib 'file::tests::minifat_stream_'
cargo test --locked -p litchi-cfb --lib 'file::tests::direct_ranges_'
cargo test --locked -p litchi-cfb --lib 'allocation_validation_tests::'
cargo test --locked -p litchi-cfb --lib 'validation_tests::fat_topology_rejection_is_a_deterministic_structured_issue'
cargo test --locked -p litchi-cfb --test sequential_writer 'mixed_minifat_fat_nested_storage_clsids_reopen'
cargo test --locked -p litchi-cfb --test sequential_writer 'difat_boundary_streams_without_retaining_source_payload'
cargo test --locked -p litchi-cfb --test sequential_writer 'exact_fat_boundary_has_no_difat_and_one_over_has_difat'
cargo test --locked -p litchi-cfb --lib
```

Retention requires those owner/consumer tests plus the matched 0511 after
profile, the tiny/few-large CFB guards, and native timing/allocation/RSS
captures. Until those artifacts exist, this document supports candidate safety
and testability only; it makes no speedup, memory, or broad compatibility
claim.

## Governing evidence

The review follows the unchanged [0511 ADR manifest](adr-manifest.json) and
the direct constraints in [ADR 0001](../../../adr/0001-priorities-and-api-layers.md),
[ADR 0005](../../../adr/0005-io-memory-and-performance.md),
[ADR 0006](../../../adr/0006-validation-security-and-compatibility.md),
[ADR 0008](../../../adr/0008-migration-and-verification.md), and
[ADR 0024](../../../adr/0024-current-topology.md). The accepted
`bf5b7f50f` exact-reservation precedent supports the local capacity pattern;
it does not replace the matched evidence required for this different decode
loop.
