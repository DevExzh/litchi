# 0779 — OPC input allocation investigation

Base: `345b81ce8a`. Scope: shared OPC owned ingress reached by ordinary XLSX
open; iWork excluded. The prior turn made verified progress (0778 integrated).
All 35 previously read architecture inputs are byte-identical to 0778.

## Corrected queue

The 0778 planning note incorrectly recommended ZIP-2 as new work. Current
`IndexedArchive::member_source` and `get_entry_from` already implement it;
0611 records its integration and 0623 adds structural prefetch. The sealed
0778 note remains historical evidence, superseded here. No duplicate ZIP
optimization is authorized by that stale recommendation.

## Current hypothesis

`Workbook::open` calls `OpcPackage::open_with_limits`, then
`read_owned_path_with_limits` and `phys_pkg::read_limited`. The latter reads
8 KiB chunks, starting with an 8 KiB reservation and requesting exact growth
for every subsequent chunk. On the retained 4,226,429-byte generated XLSX,
the sum of requested capacities is 1,092,697,469 bytes over 515 reallocations.
This is an allocation-request model, not a physical-copy or RSS measurement.
It closely matches the 0778 lifecycle's 1,100,160,722 requested bytes; independent
phase observations are required before attributing the difference.

The candidate changes capacity growth only: when a successful read needs more
capacity, grow by at least 8 KiB and otherwise one eighth of current capacity,
cap the requested capacity at the existing input limit, and reserve fallibly.
The 8 KiB I/O requests, successful bytes, interrupt retries, late I/O errors,
invalid read-count checks and exact-limit one-byte probe stay in the same order.
No metadata hint, ambient behavior, parallelism, cache or unsafe code is added.
Retained spare capacity can increase; allocation request reduction does not
prove latency or live-memory improvement. The candidate is conditional on
matched native/allocation observations and explicit memory review.

## ADR mapping

- 0001/0002/0010/0011/0024: physical ingestion remains inside OPC, public APIs
  and dependency direction unchanged.
- 0003/0006: source bytes, no-op identity and malformed-input refusals remain
  authoritative; no normalization or repair is introduced.
- 0005: input ceiling and fallible allocation remain; requested capacity never
  exceeds that ceiling. Allocator-internal overhead is not a portable bound.
  Separate native and allocator processes, raw samples, output identities and
  retained-memory tradeoffs gate any claim.
- 0008: focused boundaries and owner gates are evidence for this scope only;
  no format certification or CRUD-row promotion follows automatically.

## Planned measurement

A standalone probe copies the existing harness allocation observer and calls
only public workbook APIs. It measures open, edit, save and lifecycle separately
on the retained generated and real XLSX sources, verifying source/output
identity and the A1 marker outside each region. Save uses explicit NoSync to
study ingestion without changing ordinary Full semantics. Native timing and
allocation use separate binaries. Heaptrack is diagnostic, never native timing.

Before/after binaries must have identical probe and compiler options, with
source/binary/fixture hashes retained. Alternate process order across blocks;
retain all >5% timing/RSS flags. Phase peaks must not be added or subtracted.
No cold-cache, remote-range, concurrency or whole-program completion claim is
made by this scoped experiment.
