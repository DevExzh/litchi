# Change 0109: managed XLSX scalar-cell evidence controls

Date: 2026-08-14

The performance harness now exposes an opt-in managed source-backed tranche for
the committed XLSX value-only editor. It adds matched bounded source-backed
controls for:

- one existing cell;
- the deterministic `ceil(1%)` existing-cell set;
- the exact 256-cell batch limit; and
- a two-cell, two-worksheet transaction.

The existing eager/source-backed one-cell, `ceil(1%)`, and exact-256 controls
remain available as semantically exact eager controls. The new bounded
multi-sheet control has no eager twin because the source-backed selector is the
specific cross-worksheet closure under evidence. All selectors use the same
explicit finite source-cache entry/byte policy; managed controls additionally
construct an `ExecutionContext` and charge retained/in-flight OPC `PartData`
payload reservations to a local `Budget`.

Each measured iteration records separate open, selector-planning, commit,
stream-publication, and reopen/verification intervals. The reported elapsed
interval is the first four segments. Reopen, semantic-cell verification,
package graph/content-type/relationship checks, exact output and semantic
hashes, untouched raw ZIP-member fingerprints, cache diagnostics, source read
counters, and materialization counters are evidence gates outside that elapsed
sum. The source path uses the deterministic media-rich four-sheet
`medium`/`dense-sparse` corpora: eight 512 KiB media Parts remain untouched,
and every unselected member's local record and central record (with the offset
normalized) must remain byte-identical.

The JSON nested `source.xlsx_cell_values` object exposes the timing vectors,
source read calls/bytes, successful payload materializations, cache retention
and reservation counters, managed-budget mode/limit/use, use after the source
package is consumed, and use after all retained snapshot/commit handles are
dropped. Managed iterations require the final use to be zero; unmanaged
controls require zero budget accounting. The harness also checks deterministic
output and canonical post-edit semantic hashes across measured samples.

This is an evidence-control tranche, not a speedup result. No controlled
release ABBA comparison has been run, so it makes no latency or throughput
claim. The corpus does not provide allocation counts, RSS/peak-resident-memory
measurements, hardware/CPU pinning, cold-I/O, decompression, or real-producer
breadth evidence. Per ADR 0005 and the XLSX editor contract, the managed
`Budget` covers retained and in-flight OPC `PartData` payload reservations
only; parsed cell stores, relationship/graph metadata, staging allocations,
rewritten candidate XML, and output buffers are governed by separate bounds and
are not represented as a complete allocation or memory-accounting result.
The controls remain opt-in and do not change the default matrix; iWork remains
out of scope while the `iwa-*` crates are changing independently.
