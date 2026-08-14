# Change 0096: XLSX source-provenance publication reuse

Date: 2026-08-14

Status: Accepted for the bounded source-backed scalar-cell publisher.

## Problem

`cell_values::SourceBackedEditor` captured a complete, immutable worksheet
closure while staging a commit, but publication loaded and parsed the selected
worksheet closure again before invoking the already validating low-level OPC
overlay publisher. The second semantic parse did not strengthen the source
identity proof: every source-backed snapshot already carries the package's
unforgeable process-local `SourceLineage` and checked `SourceVersion`.

## Change

Publication now classifies the retained provenance before output:

- an exact lineage/version match continues directly to the OPC overlay
  publisher;
- a foreign lineage or stale revision refuses before output; and
- an owned snapshot without source provenance retains the previous full
  reload plus `same_source` fallback.

The OPC publisher still verifies source freshness, the selected Part's ZIP
framing, size and CRC, XML, topology, signatures, limits, raw preservation and
sink progress. No public type, dependency, cache or execution policy changed.

The regression fixture makes one selected worksheet larger than the default
8 MiB source payload cache. An independent reload demonstrably reaches a
rejected physical ZIP range, while exact no-op publication performs no such
reload and remains byte-identical. Foreign lineage, changed revision,
multi-sheet ownership and the existing lifecycle suite remain covered.

## Balanced release evidence

The same frozen binaries ran on CPU 2 in `before-A / after-A / after-B /
before-B` order, with 10 warm-ups and 100 samples per cell. The timed interval
is the existing matched harness interval: open, selector-first staging/commit
and sequential publication. Full reopen, semantic/topology checks, raw
unselected-member identity and hashes remain outside timing.

| Source-backed case | before-A -> after-A p50 | before-B -> after-B p50 |
|---|---:|---:|
| Medium one cell | 11.679 -> 8.956 ms (-23.31%) | 11.753 -> 9.056 ms (-22.95%) |
| Medium `ceil(1%)` | 43.670 -> 34.308 ms (-21.44%) | 43.447 -> 34.202 ms (-21.28%) |
| Medium exact-256 | 44.129 -> 34.662 ms (-21.45%) | 44.661 -> 34.360 ms (-23.06%) |
| Dense/sparse one cell | 78.237 -> 62.177 ms (-20.53%) | 79.377 -> 61.238 ms (-22.85%) |
| Dense/sparse `ceil(1%)` | 84.691 -> 66.556 ms (-21.41%) | 86.669 -> 66.945 ms (-22.76%) |
| Dense/sparse exact-256 | 85.009 -> 66.465 ms (-21.81%) | 86.834 -> 66.855 ms (-23.01%) |

Across the six source-backed cells, p50 geomean deltas are **-21.66%** and
**-22.65%**; p95 geomean deltas are **-21.38%** and **-22.70%**. Output hashes
match in all four legs. Physical source read and materialization counters are
unchanged (three materializations for one-cell cases and six for the other
cases), which correctly limits the claim to removal of the repeated semantic
worksheet reload/reparse rather than physical I/O or decompression.

The medium eager exact-256 after-A control is disclosed: p50 moved +30.59% and
p95 +105.28%, opposite the paired source improvement. It normalized to +1.63%
p50/+4.25% p95 in after-B. No eager-path performance claim is accepted.

Exact report and binary hashes, every p50/p95/p99 value, adjacent-direction
deltas, source/sink counters and output hashes are retained in the
[compact ABBA summary](../results/xlsx-rtf-abba-0108-summary.json).

No allocation, peak-heap, RSS, physical cold-I/O, decompression-byte or
recompression-byte conclusion is made. The capability remains limited to the
existing bounded, conservative scalar-cell closure.
