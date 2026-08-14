# Change 0097: bounded ASCII batching for streaming RTF creation

Date: 2026-08-14

Status: Accepted for `StreamingRtfWriter`'s safe ASCII emission path.

## Problem

The forward-only RTF writer already retained only 37 bytes of encoder state,
but it sent each ordinary ASCII character to the caller's `Write` sink
separately. The deterministic large streaming case therefore made 7,208,970
sink calls for a 10,092,579-byte document.

## Change

Contiguous printable ASCII that needs no RTF escaping is emitted directly in
chunks of at most 32 bytes. Backslash, braces, CR/LF, non-ASCII and malformed
or split UTF-8 retain the established scalar paths. Each candidate chunk holds
atomic hierarchical Work and Output reservations; a failed reservation falls
back to scalar emission, preserving exact limit behavior. Successful and
failed writes account for accepted input, output and attempted work exactly,
and cancellation is checked before and after every bounded sink write.

The writer gains no retained buffer: its explicit encoder state remains 37
bytes. Hard-coded pre-change wire fixtures cover plain ASCII, escapes/braces,
Windows-1252 and CR/LF. Additional tests cover parent-budget fallback,
partial/zero/interrupted/excess sink progress, cancellation during short and
full writes, malformed/split UTF-8, poisoning and semantic reopen.

## Balanced release evidence

The same CPU-2 release ABBA protocol as change 0096 used 10 warm-ups and 100
samples per cell. A non-seek hashing sink retains zero output bytes; a complete
artifact is separately reopened and checked outside timing.

| Shape | before-A -> after-A p50 | before-B -> after-B p50 | sink calls before -> after |
|---|---:|---:|---:|
| Tiny | 0.302 -> 0.071 ms (-76.61%) | 0.294 -> 0.070 ms (-76.06%) | 3,530 -> 714 |
| Medium | 38.195 -> 9.068 ms (-76.26%) | 38.487 -> 8.959 ms (-76.72%) | 450,570 -> 90,122 |
| Large | 612.059 -> 144.687 ms (-76.36%) | 612.956 -> 143.305 ms (-76.62%) | 7,208,970 -> 1,441,802 |

Across the three shapes, p50 geomean deltas are **-76.41%** and **-76.47%**;
p95 geomean deltas are **-75.23%** and **-75.76%**. Exact accepted byte counts
and SHA-256 outputs match all four legs. The largest sink request changes from
13 bytes to the deliberate 32-byte ceiling; no request can scale with caller
input size.

Exact report and binary hashes, every p50/p95/p99 value, adjacent-direction
deltas, sink counters and output hashes are retained in the
[compact ABBA summary](../results/xlsx-rtf-abba-0108-summary.json).

No allocation, peak-heap, RSS, physical cold-I/O or existing-document edit
claim is made. This result applies only to fresh forward-only RTF creation;
the separate logical-tail transaction still retains its complete validated
candidate snapshot.
