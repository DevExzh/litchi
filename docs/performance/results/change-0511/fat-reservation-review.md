# Initial push-only experiment review (rejected)

This review describes the initial push-only variant, preserved under
`initial-push/`. The final sector-batched candidate is reviewed separately.
Its references to the applied candidate refer to that initial variant.

# 0511 CFB FAT reservation candidate review

Reviewed 2026-09-11 at base revision
`b0b9ebb577d616a9a0ed93f16619a1b79f3b3bc7`. The initial review was
documentation-only; this follow-up records the owner-applied candidate in the
current [`load_fat`](../../../../crates/litchi-cfb/src/file.rs#L737) and
[`load_minifat`](../../../../crates/litchi-cfb/src/file.rs#L887) paths. The
initial [`file.rs` diff](initial-push/candidate.patch) source SHA-256 is
`414dc125d8879600d7dfe08107651ad0b39fc21454cc479d61f21417d46c4bba`, matching
the 0511 [initial source manifest](initial-push/source-manifest.json); the pre-candidate
control hash remains recorded in [before/source-manifest.json](before/source-manifest.json).
No after retention claim is made by this review.

## Disposition

The only mechanically admissible source candidate is the inner FAT-table
change at [`file.rs:857`](../../../../crates/litchi-cfb/src/file.rs#L851-L861):

```rust
try_push(&mut fat, entry, "FAT entries")?;
```

becomes `fat.push(entry)` after the existing exact fallible reservation at
[`file.rs:845-849`](../../../../crates/litchi-cfb/src/file.rs#L841-L850). This
follow-up confirms the actual diff is limited to this call and its two-line
capacity comment. The candidate is admitted as a controlled experiment;
matched after evidence and the owner/consumer validation suite still determine
retention.

The MiniFAT loop at [`file.rs:917-922`](../../../../crates/litchi-cfb/src/file.rs#L917-L922)
stays unchanged. Its count is also mathematically exact, but the retained
profile does not isolate that call site from the aggregate `try_push::<u32>`
symbol. The available evidence therefore does not meet the profile gate for a
MiniFAT candidate.

The fresh 0511 control profile now records exactly five
`SourceBackedWorkbook::from_read_at_with_limits` calls and 15,127,968 exclusive
Callgrind instruction references. `load_fat` accounts for 9.30% of that
operation-local control profile; the scope excludes selected-query, setup, and
drop work ([`admission.json`](admission.json),
[`before/callgrind-receipt.json`](before/callgrind-receipt.json)). This is
operation attribution for candidate selection, not a before/after speedup or
latency result. The admission record keeps the larger chain and ownership
validation loops mandatory and requires matched after rows before retention.

## Evidence and governing constraints

The 0511 plan identifies FAT entry construction as the open lead and binds the
review to the existing [ADR manifest](adr-manifest.json), which records all 29
numbered ADRs plus the README at the base revision. The directly applicable
constraints are:

- [ADR 0001](../../../adr/0001-priorities-and-api-layers.md): correctness,
  safety, and typed failures take priority over a local optimization.
- [ADR 0005](../../../adr/0005-io-memory-and-performance.md): a hotspot is a
  candidate signal; retention requires matched, scoped evidence and must not
  turn a descriptive profile into a speedup claim.
- [ADR 0006](../../../adr/0006-validation-security-and-compatibility.md):
  malformed-input handling and compatibility behavior remain part of the
  contract.
- [ADR 0008](../../../adr/0008-migration-and-verification.md): the source,
  correctness, and performance gates remain separately evidenced.
- [ADR 0024](../../../adr/0024-current-topology.md): CFB ownership remains in
  `litchi-cfb`; this review adds no dependency or API edge.

The retained [0413 paired profile](../change-0413/analysis/profile.json) is
descriptive `cycles:u` attribution from 10,000 samples on CPU 2. In its
candidate whole-process ranking, `load_fat` has 6.2418% inclusive presence and
`try_push::<u32>` has 3.9850%; the source-backed leaf ranking gives 5.1759% and
3.9676%, respectively. Those values identify work worth isolating. They do not
measure this proposed one-line change, and `try_push::<u32>` combines FAT
entries, MiniFAT entries, sector-location vectors, and other `u32` call sites.
The profile explicitly makes no phase-latency or speedup claim, so it supports
the FAT-only investigation but cannot support changing MiniFAT.

The accepted precedent is commit `bf5b7f50f`
(`perf(cfb): use established exact chain scratch reservation`). It retains a
fallible exact reservation and replaces the corresponding exactly bounded
`try_push` with `push`, while preserving the established allocation boundary.
The applied FAT edit has the same shape, but still requires its own matched
after measurement and validation run.

## FAT capacity proof

The current path establishes the following facts before the candidate loop.

1. `open_with_limits` accepts only CFB version 3 or 4 sector shifts, yielding a
   sector size of 512 or 4096 bytes; both are divisible by four. The private
   loader repeats the physical-count checks before any table allocation.
2. Header and DIFAT traversal fills `fat_sectors` and then checks
   `fat_sectors.len() == expected_fat_sectors` at
   [`file.rs:834-839`](../../../../crates/litchi-cfb/src/file.rs#L834-L839).
   The variable-length list construction keeps `try_push` because malformed
   header/DIFAT input can stop before the declared count.
3. The table vector is created empty by
   [`try_vec_with_capacity`](../../../../crates/litchi-cfb/src/file.rs#L2486-L2491)
   with
   `fat_entry_count = fat_sectors.len() * (self.sector_size / 4)` after a
   checked multiplication at [`file.rs:842-849`](../../../../crates/litchi-cfb/src/file.rs#L841-L850).
   A successful `try_reserve_exact` guarantees capacity at least that count.
4. The outer loop visits every FAT sector exactly once. For each sector,
   `sector_data.as_chunks::<4>().0` yields exactly `self.sector_size / 4`
   four-byte chunks. The inner loop therefore performs exactly
   `fat_entry_count` pushes, and no other code pushes into `fat` before or
   during this loop.
5. Consequently, immediately before each candidate `fat.push(entry)`,
   `fat.len() < fat_entry_count <= fat.capacity()`. The push cannot request a
   reallocation. The vector may have spare capacity because the allocator is
   allowed to reserve more than requested, which only strengthens the proof.

The proof depends on the count check and exact reservation remaining adjacent
to this loop. Any future change that adds a push, changes sector geometry, or
allows a partial-chunk interpretation must restore a fallible growth check.

## Error order and preservation proof

Replacing only the inner FAT-table `try_push` preserves the observable format
and I/O sequence:

- Count conversion, physical-file bounds checks, header/DIFAT reads, sector
  claims, and DIFAT marker validation remain before the exact FAT-entry
  reservation.
- The `FAT entries` exact reservation remains fallible, with the same resource
  label and at the same point before the first FAT-sector read. An allocation
  failure at that boundary remains a typed `OleError::Allocation`.
- Each FAT sector is still read in the same order by `read_sector_into`, and
  each four-byte entry is still decoded by `read_u32_le` before it is inserted.
  Truncated final sectors therefore retain the existing zero-fill behavior and
  the same parse order.
- FAT-sector and DIFAT-sector ownership-marker checks remain after all table
  entries are decoded, and `self.fat` is still published only after those
  checks succeed. The resulting table values and publication timing are
  unchanged.
- The surrounding `fat_sectors`, `difat_sectors`, chain, directory, and
  validation vectors retain `try_push`. Their lengths are data-dependent or
  walked through malformed input, so this proof does not generalize to them.

`try_push` itself performs `try_reserve(1)` and then `push`
([`file.rs:2494-2500`](../../../../crates/litchi-cfb/src/file.rs#L2494-L2500)).
After the successful exact reservation, every FAT-table call has one unit of
capacity available; under the `Vec` capacity contract, those repeated
reservations are no-ops. Removing them changes only the redundant allocation
probe. A successful reserve within existing capacity makes no allocator call.
Removing that check therefore does not remove an allocator-visible failure
point or change corruption, I/O, marker, or publication error ordering.

## Why MiniFAT is deferred

`load_minifat` obtains an exact chain from
[`collect_sector_chain_exact`](../../../../crates/litchi-cfb/src/file.rs#L2660-L2704),
checks `sectors.len() * self.sector_size` with checked multiplication, and
reserves `entries_count = data_len / 4` before its entry loop. Thus the same
arithmetic would make its pushes capacity-safe for valid CFB geometry.
However, the retained 0413 profile reports one aggregate `try_push::<u32>`
symbol. It cannot distinguish the MiniFAT loop from FAT entries, FAT/DIFAT
locations, or other `u32` uses. No MiniFAT-specific profile contribution is
therefore measured. Keep `try_push` at `file.rs:921` until an isolated
MiniFAT workload and profile establish a material contribution and a separate
candidate passes the same malformed-input and preservation gates.

## Existing focused coverage

The following tests were inspected as the regression surface. They are existing
tests; this review does not add or execute them.

**FAT/DIFAT ownership, malformed metadata, and physical bounds**

- `crates/litchi-cfb/src/file.rs`: `malformed_large_declarations_do_not_unwind`,
  `rejects_self_referential_difat_chains`,
  `rejects_difat_counts_beyond_the_physical_file_before_reserving`,
  `tolerates_nonfree_fat_padding_beyond_the_physical_file`,
  `rejects_sector_reads_past_the_physical_file`.
- `crates/litchi-cfb/src/allocation_validation_tests.rs`:
  `hostile_count_metadata_is_rejected_before_index_allocation`,
  `hostile_difat_and_minifat_counts_are_rejected_before_index_allocation`,
  `version_4_directory_byte_limit_is_checked_before_chain_allocation`,
  `rejects_difat_start_count_disagreement`,
  `rejects_directory_and_fat_sector_overlap`,
  `rejects_fat_sector_without_fatsect_marker`,
  `rejects_incorrect_minifat_sector_count`,
  `still_rejects_a_header_difat_list_shorter_than_its_count`, and
  `still_rejects_sectors_that_start_past_the_end_of_the_file`.
- `crates/litchi-cfb/src/validation_tests.rs`:
  `fat_topology_rejection_is_a_deterministic_structured_issue`.

**FAT/MiniFAT chain length, marker, cycle, and truncation behavior**

- `crates/litchi-cfb/src/file.rs`: `rejects_chain_lengths_before_reserving_chain_storage`,
  `detects_cycles_in_allocation_chains`,
  `fat_stream_reads_only_the_declared_logical_size`,
  `fat_stream_replays_fragmented_chain_in_order_at_exact_size`,
  `fat_stream_chain_errors_remain_typed_and_ordered`,
  `fat_stream_read_zero_fills_a_truncated_final_sector`,
  `minifat_stream_replays_fragmented_chain_in_order_at_exact_size`,
  `minifat_stream_chain_errors_remain_typed_and_ordered`,
  `direct_ranges_follow_fragmented_fat_and_minifat_chains`,
  `direct_ranges_cover_boundaries_and_preserve_truncated_sector_zero_fill`,
  `direct_ranges_reject_bad_metadata_without_unbounded_traversal`,
  `direct_ranges_validate_bounds_before_payload_io_and_allow_exact_eof`,
  `read_sector_into_zero_fills_a_truncated_final_sector`,
  `rejects_minifat_buffer_size_overflow_before_reading_a_sector`, and
  `reports_minifat_allocation_failure_without_panicking`.
- `crates/litchi-cfb/src/allocation_validation_tests.rs`:
  `open_rejects_minifat_cycles_and_invalid_markers`,
  `open_rejects_fat_stream_cycles_and_invalid_markers`,
  `open_rejects_fat_stream_short_excess_and_terminal_markers`, and
  `open_rejects_minifat_stream_short_excess_and_terminal_markers`.

**Writer-produced FAT/DIFAT/MiniFAT reopen coverage**

- `crates/litchi-cfb/tests/sequential_writer.rs`:
  `mixed_minifat_fat_nested_storage_clsids_reopen`,
  `difat_boundary_streams_without_retaining_source_payload`, and
  `exact_fat_boundary_has_no_difat_and_one_over_has_difat`.

## Exact scoped commands for a future candidate run

These commands are the narrow validation set for the one-line FAT candidate;
they were recorded rather than run during this documentation review:

```text
cargo test --locked -p litchi-cfb --lib 'file::tests::fat_stream_'
cargo test --locked -p litchi-cfb --lib 'file::tests::minifat_stream_'
cargo test --locked -p litchi-cfb --lib 'file::tests::direct_ranges_'
cargo test --locked -p litchi-cfb --lib 'file::tests::rejects_'
cargo test --locked -p litchi-cfb --lib 'file::tests::detects_cycles_in_allocation_chains'
cargo test --locked -p litchi-cfb --lib 'file::tests::tolerates_nonfree_fat_padding_beyond_the_physical_file'
cargo test --locked -p litchi-cfb --lib 'allocation_validation_tests::'
cargo test --locked -p litchi-cfb --lib 'validation_tests::fat_topology_rejection_is_a_deterministic_structured_issue'
cargo test --locked -p litchi-cfb --test sequential_writer 'mixed_minifat_fat_nested_storage_clsids_reopen'
cargo test --locked -p litchi-cfb --test sequential_writer 'difat_boundary_streams_without_retaining_source_payload'
cargo test --locked -p litchi-cfb --test sequential_writer 'exact_fat_boundary_has_no_difat_and_one_over_has_difat'
cargo test --locked -p litchi-cfb --lib
```

After a candidate source edit, the existing 0511 control/after harness should
also run its already-defined build, test, and evidence lanes. This review does
not claim those lanes passed and does not change the running control build.
