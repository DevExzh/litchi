# Change 0511 CFB source-open residual review

Status: this review selected the initial push-only experiment, subsequently rejected
on instruction attribution. The final sector extension and capacity proof are in
[sector-batch-review.md](sector-batch-review.md); measured results and gate
disposition are in [0511](../../changes/0511-cfb-fat-entry-reservation.md).
The proposal and suggested checks below are retained as investigation history.
Existing v3/v4 malformed-container and source-fence tests cover the final local
proof; no duplicate implementation-mirroring differential test was added.

Scope: CFB/OLE2 source opening and validation only. This is a read-only source
review; it makes no production change and reports no new measurement.

Reviewed revision: `b0b9ebb577d616a9a0ed93f16619a1b79f3b3bc7` (`b0b9ebb57`).
The review is anchored on [0413 CFB chain scratch reservation](../../changes/0413-cfb-chain-scratch-reservation.md)
and the CFB portion of the [0510 OLE2/OOXML priority review](../change-0510/ole2-ooxml-priority-review.md).

## Initial investigation decision

The next bounded experiment should measure one call site in
`OleFile::load_fat`: replace the per-entry `try_push` at
[file.rs#L849](../../../../crates/litchi-cfb/src/file.rs#L849)
through [file.rs#L857](../../../../crates/litchi-cfb/src/file.rs#L857)
with `fat.push(entry)` after the existing exact fallible reservation. No
candidate is admitted by this review and no end-to-end speedup is claimed.

This is the smallest remaining version of the 0413 pattern. It removes a
fallible-reserve check from the hot FAT-entry loop while retaining the existing
reservation, source reads, decoding, marker checks, ownership checks, and all
resource limits. The header and DIFAT `try_push` calls, MiniFAT entry loop,
directory entry loop, and exact chain collector should stay out of the first
experiment so their error and allocation behavior can be attributed
independently.

## Current CFB open and validation path

`SharedOleFile::open_source_with_limits` obtains the source version and length,
applies the input limit, and passes a positional `ReadAtCursor` to the
canonical `OleFile::open_with_limits` parser. It checks the source version again
after parsing before retaining the parsed index; see
[shared.rs#L479](../../../../crates/litchi-cfb/src/shared.rs#L479)
through [shared.rs#L526](../../../../crates/litchi-cfb/src/shared.rs#L526).
Those source fences and the `ReadAt` boundary are part of the CFB contract.

`OleFile::open_with_limits` checks the header and physical sector count, loads
the FAT and directory, optionally loads the MiniFAT, validates every stream
allocation, and reconciles the physical sector layout before publishing the
index. The order is visible at
[file.rs#L539](../../../../crates/litchi-cfb/src/file.rs#L539)
through [file.rs#L720](../../../../crates/litchi-cfb/src/file.rs#L720).
The repeated count checks before allocation are deliberate hostile-input
guards; the candidate must not move or merge them.

Validation deliberately reparses structural sectors through the same canonical
open path. `validate_source_with_limits` and `SharedOleFile::validate` retain
the pre/post source fences and report deterministic structural issues while
leaving source I/O, instability, and allocation failures as errors; see
[validation.rs#L89](../../../../crates/litchi-cfb/src/validation.rs#L89)
through [validation.rs#L150](../../../../crates/litchi-cfb/src/validation.rs#L150).
Reusing an already retained index to avoid that reparse would require a new
validation contract and is outside this change.

## Candidate: direct push for the already-sized FAT vector

The `load_fat` path first collects the declared FAT sector locations and checks
that the final list has exactly the declared count at
[file.rs#L791](../../../../crates/litchi-cfb/src/file.rs#L791)
through [file.rs#L839](../../../../crates/litchi-cfb/src/file.rs#L839).
It then computes `fat_entry_count` with checked arithmetic and performs an
exact fallible `try_vec_with_capacity` before reading FAT sector contents. The
payload loop currently does this for every decoded `u32`:

~~~text
read_u32_le(chunk, ...)?;
try_push(&mut fat, entry, "FAT entries")?;
~~~

The relevant reservation and loop are at
[file.rs#L842](../../../../crates/litchi-cfb/src/file.rs#L842)
through [file.rs#L857](../../../../crates/litchi-cfb/src/file.rs#L857).
After the count and sector-list checks, every FAT sector contributes exactly
`sector_size / 4` full chunks. The validated sector size is 512 or 4096 bytes,
and the number of sectors is the checked declared count, so the loop executes
exactly `fat_entry_count` pushes. The vector therefore has enough capacity for
every entry before the first source read. Replacing only the second line with
`fat.push(entry)` cannot grow the vector on the current path.

The exact reservation remains the sole allocation point and keeps its current
`"FAT entries"` resource label. `read_u32_le` still precedes insertion, and
the marker validation loop still runs after all entries are decoded. A short
physical final sector is still zero-filled by the existing sector-read path and
is handled by the same `as_chunks::<4>().0` behavior over the fixed sector
buffer. There is no source-read, FAT
shape, ownership, or limit change in this candidate.

The safety condition is local and should remain explicit in any implementation:
the direct push is valid only after the exact capacity reservation and the
exact `fat_sectors.len()` check. A future refactor that weakens either proof
must restore a fallible push or add an equivalent checked assertion; do not
change the generic `try_push` helper, which remains necessary for dynamic
collections.

## Why this is the first profile target

The 0413 whole-process profile, inclusive of source open, still showed
`try_push::<u32>` at 3.985% of candidate cycles and `load_fat` at 6.242% after
the scratch collector optimization. The corresponding pre-0413 control showed
8.338% and 1.721%; these shares are diagnostic and are affected by inclusive
call attribution, so they are not a speedup estimate. They do identify the
remaining FAT-entry helper as a measurable operation target.

The fresh 0511 control is a tighter source-open workload: five
`SourceBackedWorkbook::from_read_at_with_limits` calls with setup excluded from
the bracket. Its 15,127,968 instruction references are distributed as follows
in the [exclusive callgrind report](before/callgrind-exclusive.txt):

| CFB operation | Exclusive Ir |
| --- | ---: |
| `SectorChainScratch::collect_exact` | 37.03% |
| `OleFile::claim_sector` | 16.46% |
| `validate_physical_sector_layout` | 13.17% |
| `validate_stream_allocations` | 13.08% |
| `load_fat` | 9.30% |

The larger loops are not a clearly safe replacement for the FAT push. The
scratch collector walks each declared stream chain with cycle, terminal-marker,
length, and limit checks; `claim_sector` records ownership and rejects overlap;
`validate_physical_sector_layout` reconciles every physical sector; and
`validate_stream_allocations` applies those checks to every directory stream.
Sharing or removing any of those walks would need a new bounded ownership/index
representation and an exact error-order proof. The current source does not
provide one, and the validation ADRs require these structural checks to remain.
They should be separate investigations even though they dominate this control
profile. The 9.30% `load_fat` share keeps the one-line direct-push candidate
worth measuring without treating the larger validation loops as redundant.

The fresh profile should use an operation-local source-open bracket and compare
the unchanged control with the single-call-site candidate on:

1. A plain owned CFB/XLS source with a FAT-heavy layout, where
   `OleFile::load_fat` and the entry loop have enough iterations to expose the
   eliminated reserve check.
2. A metadata-only open and an eager stream workflow, because 0413 found that
   source-open work is diluted differently by downstream consumers. Report
   `load_fat`/entry-loop attribution where the profiler can distinguish the
   call site, plus end-to-end source-open p50, p95, and p99.
3. Small CFBs, MiniFAT-bearing files, and the 0413 few-large regular-FAT guard.
   The guard is required because a local improvement must not regress the CFB
   shapes that previously showed roughly 2.8–3.1% slower behavior.

Capture allocation calls/bytes and peak RSS alongside source read calls/bytes;
the direct push is expected to change neither retained capacity nor configured
limits. Cycles/instructions, branches, and branch misses are useful when PMU
access is available. Whole-process profiles may remain diagnostic, but a
release claim needs the operation-local timings and the existing >5% latency or
RSS review trigger. No memory-reduction claim follows from this candidate.

## Correctness and test requirements if implemented

The existing CFB tests cover the invariants that make the direct push safe and
must remain green, especially the allocation and malformed-input cases in
[allocation_validation_tests.rs#L183](../../../../crates/litchi-cfb/src/allocation_validation_tests.rs#L183)
and the CFB validation report/error-order cases in
[validation_tests.rs#L180](../../../../crates/litchi-cfb/src/validation_tests.rs#L180).
The `load_fat` malformed DIFAT/self-reference tests and FAT/MiniFAT allocation
tests in [file.rs#L3830](../../../../crates/litchi-cfb/src/file.rs#L3830)
should be included in the focused run.

If the candidate is coded, add a focused differential check using writer-built
CFBs with both 512-byte and 4096-byte sectors, then mutate the same fixtures to
cover:

- declared FAT/DIFAT count boundaries and hostile counts;
- wrong FAT markers, overlaps, cycles, invalid DIFAT references, and a
  truncated final physical sector;
- non-FREESECT padding and unrelated sectors that exercise the physical layout
  reconciliation;
- source I/O and version changes at the existing fences.

Compare the candidate with the ordered helper behavior, including the exact
`OleError` display and typed error where applicable. The test must establish
that the early reservation failure remains the same resource failure and that
no newly reachable path performs an infallible growth. Root should run the full
CFB/XLS/DOC/PPT feature matrix, formatting, lint, documentation, and performance
guards; this review intentionally runs none of them.

## Deferred CFB work

`load_directory` first builds `ValidatedDirectoryEntry` values and validates the
complete sibling topology, then `build_storage_tree_iterative` parses reachable
entries again into public `DirectoryEntry` values; see
[file.rs#L960](../../../../crates/litchi-cfb/src/file.rs#L960)
through [file.rs#L980](../../../../crates/litchi-cfb/src/file.rs#L980) and
[file.rs#L1474](../../../../crates/litchi-cfb/src/file.rs#L1474)
through [file.rs#L1554](../../../../crates/litchi-cfb/src/file.rs#L1554).
This is a real repeated walk, but 0413 inclusive attribution was only about
0.89% for `validated_directory_entries` and 1.45% for `load_directory` in the
candidate. Combining the passes would also risk Root Entry normalization,
historical public decoding, CLSID formatting, SID-aligned name-cache identity,
tree error order, and bounded traversal behavior. It needs a separate
operation-local profile and an exact differential test.

Two other exact-reservation sites resemble the FAT loop but should be measured
separately: MiniFAT payload insertion at `file.rs:910–921` and the generic exact
chain collector at `file.rs:2660–2728` (its `try_push` is at line 2703). Their
callers have different truncation and error-order contexts. The 0413
`SectorChainScratch::collect_exact` direct-push change is the reference pattern,
not permission to batch these sites into 0511.

## References

- [0413 change record](../../changes/0413-cfb-chain-scratch-reservation.md)
- [0413 results](../change-0413/README.md)
- [0510 OLE2/OOXML priority review](../change-0510/ole2-ooxml-priority-review.md)
- [CFB allocation validation tests](../../../../crates/litchi-cfb/src/allocation_validation_tests.rs)
- [CFB validation report tests](../../../../crates/litchi-cfb/src/validation_tests.rs)
