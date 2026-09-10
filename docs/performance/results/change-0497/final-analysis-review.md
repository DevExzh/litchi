# 0497 formal1 capture and analysis review

The reviewed continuation completed the frozen formal1 suffix without
restarting the 191 children with existing terminal receipts (ordinals 0--190).
The resume
receipt reports 97/97 suffix children passing (`17:45:32.243154Z` through
`17:58:41.367135Z`, about 13m 09s). The original ordinal-191 interruption is
retained under `interrupted/formal1-enospc`; its child exit remains unknown and
its raw files are archived as `.raw` files. It was not converted into a
measurement or terminal receipt.

The formal namespace now contains all 288 frozen child directories. Each child
has one `started.json`, `terminal.json`, `report.json`, `resource.txt`,
`stdout.txt`, `stderr.txt`, and `replay-cleanup.json`; each of the 288 replay
cleanup receipts reports `status: pass`. The formal and pilot private scratch
roots are absent. The frozen collector wrote:

- `verification/formal1.json`: 288 children and 8,640 measured samples,
  `status: pass`.
- `analysis/formal1.json`: the regenerated analysis for the same 288 children
  and 8,640 samples.

The matrix remains 72 before/default hashing children, 72 after/default
hashing children, 72 after/counting-sink children, and 72 after/atomic-path
children. Every child has three warmups and 30 measured samples. The analysis
keeps the two default hashing repeats as the only before/after comparison. It
keeps counting and atomic publication as after-only capability records and
does not claim an atomic speedup.

All 144 after-only capability records have a production candidate artifact
that matches its route oracle. Counting uses the timed production artifact and
an untimed candidate oracle; atomic uses the timed production artifact and its
post-timer path oracle. Atomic write-call and sink-digest values remain
unobservable and are not synthesized. Among the 144 after-only capability
children, allocator metrics are available for all 72 allocator children and
are explicitly unavailable for the 72 normal-role children; normal-role RSS
remains one whole-child observation.

The default hashing comparison retains 864 metric summaries. The analysis
marks 100 summaries at or above the 5% adverse threshold: 72 allocator live
endpoint summaries, 16 latency summaries, and 12 RSS summaries. The flags are
descriptive repeat-block observations, not a causal regression claim. Across
the 18 arm-level comparisons, the median relative deltas are:

- normal latency p50/p95/p99: +1.2007% / +0.9334% / +0.2859%;
- normal whole-child RSS: +0.4109% at p50, p95, and p99;
- allocator latency p50/p95/p99: +0.2357% / +0.5677% / +0.6991%;
- allocator whole-child RSS: +1.3846% at p50, p95, and p99.

Allocator call counts and total allocated/deallocated bytes have zero median
delta across the 18 arms. The allocator live-byte endpoint shows a repeated
candidate increase of about 12.5 KiB (roughly 13.5% at the smaller workloads);
this is retained as an observed resource difference rather than hidden by the
latency summary. Tail flags include the recorded file-store and deterministic
latency arms; the complete per-arm and per-percentile values remain in
`analysis/formal1.json`.

The retained `capture-cleanup-overlap.json` records shared-disk cleanup during
the whole-child intervals for
`r2-after-hashing_sink-allocator-deterministic-latency-s64-a16384-short-c64`
and
`r2-before-hashing_sink-allocator-deterministic-latency-s64-a16384-short-c64`.
It records 16,923,615,232 removed bytes. The interval cannot attribute
writeback interference to individual samples, so the formal results are
shared-host descriptive evidence and the affected outliers receive no causal
interpretation.

The formal capture, collector, verification, analysis, and replay cleanup
gates are complete. The remaining custody gate is the final canonical cleanup
and seal receipt. The optional syscall attempts remain diagnostic or typed
host-filter failures as previously recorded; they do not block the formal
matrix and do not add operation-attribution evidence. The result also does
not establish Windows same-path or hardlink behavior, crash durability, cold
filesystem behavior, native Word producer round trips, or broad CRUD
coverage.
