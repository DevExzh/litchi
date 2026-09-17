# 0668: XLS query residues remove packed-cell records and locator scratch

Status: retained, implemented in `litchi-xls`. `performance_claim: none` —
this record reports deterministic allocation and corpus-index evidence, but no
wall-clock result or claim-registry entry.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred
until that goal completes; iWork is excluded.

This is the follow-on for rows 11 and 12 of
[0651](0651-queue-refresh-after-the-second-wave.md), after the decisions in
[0652](0652-owner-decisions-for-the-third-wave.md). It keeps the retained XLS
index and the framing fusion behind their existing gates, and removes the
proven work that can be removed without either design.

## What changed

`MulRk` and `MulBlank` worksheet records now have measure-only visitors. The
packed range and complete payload are validated first, every packed cell still
validates its XF index, and a selected target is materialized exactly as the
old path did. An unwanted `MulRk` cell no longer constructs an 88-byte
`CellRecord` or decodes its RK value; an unwanted `MulBlank` cell only passes
its measured `(row, column, XF)` identity to the sink. Whole-sheet walks still
materialize every cell.

The source-backed shared-string locator scan now gives its cursor the existing
`RecordRef` slice directly. It still walks the complete `SST`/`Continue` chain,
records the same source and logical offsets, and applies the same validation
order, but it no longer allocates a temporary `Vec<&[u8]>` containing one fat
pointer per BIFF record. The scan remains eager at open; this change does not
introduce lazy locators or retain string payloads.

## What stayed deferred

The retained cross-query sheet index and the snapshot-scoped chain hint remain
deferred. ADR 0005 requires retained clean values to live in a bounded,
weighted, evictable cache with accounting and observable eviction behavior; no
crate currently supplies that cache, and `StreamChainHint` borrows its reader.
This change does not bypass that gate with an unbounded map or a reader-free
position whose semantics have not been decided.

The two XLS framing passes also remain deferred. Change 0633 priced each at
about 1.6% of the relevant open and froze fusion because it could change the
coverage proof and refusal order; materializing all frames costs its own
memory. No framing order or refusal boundary moved here.

## Correctness and contract

The packed visitors use the existing `packed_cell_range` checks and fixed-width
reads, so malformed lengths, cell counts, and column ranges are refused before
the sink sees a cell. The selected path computes the RK value only after the
target coordinates match; all other packed cells still reach the sink's XF
validation. The whole-sheet path continues to request every cell and therefore
keeps its existing materialized output.

The direct `RecordRef` cursor uses the same payload bytes as the removed slice
vector. The pinned SST index digest remains `0x9cb14f5daa02eebc`; the corpus
differential indexes 117 fixtures and refuses the same four malformed SSTs.
There is no public API, limit, error, fence, output-byte, dependency, or
`unsafe` change.

## Evidence

The 0641 census contains 127,072 cell values: 33,758 `MulBlank` packed cells
and 26,817 `MulRk` packed cells, 60,575 values in total (47.7%). The new
measure-only route covers those unwanted query cells without making the record
that the queue identified.

The existing test-only counting allocator was run on the base `5fa92d7ce` and
on this branch with `CARGO_BUILD_JOBS=2`, `CARGO_PROFILE_DEV_DEBUG=0`, and
`CARGO_PROFILE_TEST_DEBUG=0`:

| source-backed open | base | 0668 | exact delta |
| --- | ---: | ---: | ---: |
| `54016.xls` (28 SST records) | 246 allocations / 754,968 bytes | 245 / 754,520 | −1 / −448 |
| `WithCustomViews.xls` (13 SST records) | 208 allocations / 153,205 bytes | 207 / 152,997 | −1 / −208 |

The removed bytes are exactly one `Vec<&[u8]>` reservation per scan: a
`&[u8]` is a 16-byte fat pointer on this target, so the two record counts give
the measured lower bound of 448 and 208 bytes. The 54016 whole-sheet walk is
unchanged at 35,793 allocations and 1,389,460 bytes, showing that the
resolver's retained-table path did not gain scratch work.

No timing floor was collected for this small open-time allocation change, so no
latency or speedup is claimed. The deterministic allocation deltas, the pinned
index, the packed-cell differential, and the complete package tests are the
accepted evidence.

## Validation

The branch passes `cargo fmt --all -- --check`, `cargo check -p litchi-xls`,
the focused packed-cell and SST corpus tests, the four-test
`sst_scan_allocations` integration test, and the complete `cargo test
-p litchi-xls` suite (1,068 unit tests plus all integration and doc tests).

## Breaking changes

None.

## Limitations and withheld results

No release timing, cycles, instructions, RSS, cold-cache, physical-I/O,
concurrency, or cross-platform measurement was run. The source open still
scans the entire string table to build locators, and the implementation makes
no claim for a retained cross-query cache or fused framing. Those items remain
in the queue with their 0005/0633 admission conditions.

See the [evidence packet](results/change-0668/README.md) and its
[four log sections](results/change-0668/log-sections.md).
