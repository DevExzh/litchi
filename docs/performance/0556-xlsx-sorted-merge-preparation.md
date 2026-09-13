# 0556 XLSX sorted provenance merge preparation

A checked linear merge candidate is implemented and retained as an isolated
patch for `Store::merge_omitted_cells`. It replaces append-then-sort while
preserving the ordinary constructor's validation and shared index/extent
builder. Complete XML validation, reduced parsing, fallback, and parsed
worksheet structure remain authoritative.

The baseline passed 1,313 XLSX tests; the candidate passed 1,316. Both passed
warning-denied Clippy, formatting, workspace feature checks, XLSX rustdoc, and
minimal-feature checks. Added tests compare every Stored field and structural
record with the original merge path, including distinct parsed/source metadata,
merge lookup, traversal, empty parsed input, grid edges, and refusals.

Review caught a parsed-iterator lifetime that could overlap index rebuilding;
the candidate explicitly releases that iterator before rebuilding. It still
retains the original parsed allocation during the interleaved merge, so fresh
peak-memory evidence is required. A formatting failure and a subsequent rerun
rejected by the source-integrity guard are retained; the successful corrected
runs are separate.

Production source is restored. This is preparation for a measured optimization,
with no latency, allocation, or instruction improvement claimed. The prospective
plan covers four existing XLSX shapes, native/allocator ABBA runs, commit-core
allocation scope, heaptrack call stacks, and positive Callgrind attribution.
Noise thresholds and remaining instrumentation decisions must be frozen in the
next measurement batch before candidate captures.

[Candidate, reviews, commands, and evidence](results/change-0556/README.md).
OLE2 and OOXML remain the priority; ODF is deferred until that goal completes.
