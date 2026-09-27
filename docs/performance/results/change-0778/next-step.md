# Optional next step after 0778

Status: planning note only. This file adds no performance result and recommends
no change to the default `Durability::Full` policy.

## What the allocation row says

The generated XLSX lifecycle row records 29,845 allocation calls,
5,558 reallocations, and 1,100,160,722 cumulative allocated bytes for a
4,226,568-byte published archive. Its operation-region peak live allocation is
about 14,977,463 bytes; `live_bytes_after` is about 13,542,317 bytes. The other
three durability policies have the same allocation calls, reallocations and
cumulative allocated bytes. Their absolute region-peak values carry
deterministic entry-state offsets relative to the corresponding default row:
`Full` +63 bytes, `FileOnly` +78 bytes and `NoSync` +72 bytes. After subtracting
each row's `live_bytes_before`, the peak above entry and the net live delta are
equal across all four policies; these offsets are not run noise. See
[`allocation-r1-generated-xlsx-lifecycle-full.json`](allocation-0/allocation-r1-generated-xlsx-lifecycle-full.json)
and the corresponding `default`, `file-only` and `no-sync` rows.

`allocated_bytes` is cumulative allocator traffic. The allocator observer in
`tools/perf-baseline/src/allocation_metrics.rs:661-725` adds the size of each
successful allocation and the *new* size of each successful reallocation;
the allocator counter does not report a copy volume. Repeated buffer growth can
therefore count capacity from multiple generations. The counter's scope is the
global system allocator, and its region peak is a live-byte high-water
observation, not RSS.

The lifecycle region also explains why this is not a save-only number. In
[`ordinary_save.rs`](../../../../tools/perf-baseline/src/ordinary_save.rs),
lines 1392-1424, the region and clock begin before `Owner::open`, then cover
`open`, `edit` and `save_at`; the region closes while the owner is still alive.
Readback, digest and cleanup occur afterward. The row consequently combines
package opening and model/edit allocations with publication allocations. The
XLSX package routes `save_with_durability` and `write_to` through the same
writer (`crates/litchi-xlsx/src/package.rs:1091-1147`), while
[`atomic.rs`](../../../../crates/litchi-opc/src/atomic.rs) lines 241-255 select
only the file and parent synchronization calls after the staged write. Equal
allocator rows across policies are therefore expected.

The packet's byte split is derived from archive contents and explicitly does not
observe copy-through or recompression. Its process counters also mark
compressed, decompressed and recompressed byte boundaries unavailable for this
atomic operation. The 1.1 GB figure must not be read as RSS, bytes copied,
decompressed bytes or proof that every source member was recompressed. There is
no source-grounded allocation-site attribution here that would justify changing
the XLSX writer based on this total.

## Candidate comparison

| Candidate | Evidence and reach | Decision |
| --- | --- | --- |
| Durability policy follow-up | On the generated XLSX lifecycle matrix, `NoSync` p50 is 4.812 ms and `Full` p50 is 14.265 ms. This is a large observed opt-in policy difference, while allocations stay equal; it does not optimize the open/edit/writer path. | Keep the explicit policy and its measured tradeoff. A further batch would need a separately justified cold, existing-destination or device matrix. Do not infer a default-policy change. |
| ZIP first-read span coalescing (0587 ZIP-2) | The retained source analysis (0587, lines 382-400) measured 87 positional requests for a 132-member workbook open and identifies the first member read as local header plus payload plus descriptor. A bounded central-directory span was modelled as 87 → 45 open requests, with a 2–3 request first read becoming one; the delayed range-source selectors make this relevant to request latency. | **Recommended next production batch.** Freeze the fallible span ceiling and mismatch fallback, implement the existing selector seam, and gate it with the 0582 differential plus delayed range-source request/error/output evidence. No local-file speedup should be claimed. |
| XLSX cumulative allocation reduction | The lifecycle counter mixes open, edit and save and counts allocator traffic, with no owner or phase attribution. The writer source supports streaming publication and can take preservation paths, so the total alone does not identify a safe allocation removal. | Do not select from this row. A later phase-specific attribution measurement must establish a target first. |

ZIP-2 has the better next-step ROI because it addresses the still-uncovered
range-source dimension with an existing request-count model and selectors. It is
separate from the local durability result: the latter already supplies an
owner-authorized policy choice, while ZIP-2 can reduce delayed-source request
round trips without weakening crash guarantees. The bounded implementation must
retain ADR 0005 limits, preserve per-member error identity, fall back when local
header lengths do not match the central record, and pass the 0582 differential
before any latency claim is admitted. The source record remains
[`0587-remaining-opportunity-survey.md`](../../0587-remaining-opportunity-survey.md);
its model is not a measurement on a local file.
