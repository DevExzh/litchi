# Route analysis method

`analyze_routes.py` consumes only the frozen `route-protocol.json` and the
formal `formal1` receipt trees. It requires the 120 initial route processes and
108 deterministic one-factor axis processes exactly once, with both `normal`
and `allocator` roles and repeats 1 and 2. Pilot directories are deliberately
excluded from the input and never contribute to a statistic.

Before reading a report, the analyzer verifies the terminal receipt, its
`started.json` identity, all four artifact hashes, the protocol hash, machine
and environment bindings, the copied build and build gate, the matched source
manifest, both formal capture gate receipts, and
`route-attempts/<attempt>/execution-inputs.json`. The execution-input receipt
also binds the preparation receipt, every exported archive and hash manifest,
the three staged file inputs, and workspace/evidence/scratch device and mount
observations. File-input axis receipts must match that prepared byte/hash
identity exactly. A missing, failed, replaced, or unbound receipt stops the
analysis.

Each process keeps its 30 measured samples. The report contains p50, p95, and
p99 for elapsed time, candidate throughput, source throughput, authored
throughput, process RSS delta/peak when procfs is available, and whole-process
GNU `time` maximum RSS. Allocator counters are reported in a separate object
for allocator processes and are never folded into RSS. That object keeps
allocation/deallocation/reallocation calls and bytes, live and absolute peak
counters, and the per-sample `operation_peak_increment_bytes` derived as
`region_peak_live_bytes - live_bytes_before`; its percentiles are calculated
from that per-sample vector. Source and authored throughput use the retained
source archive bytes and authored encoded XML bytes, respectively; the
candidate archive throughput is a third metric.

Each process summary also retains source calls/requested/returned bytes, sink
write calls/accepted bytes/largest write, authored opens, and replay opens/read
calls/returned bytes. File replay write calls are emitted only when the report
has a non-null file counter; null or absent counters are omitted rather than
converted to zero. Fixed six-bin source, sink, and replay histograms are kept
as an exact common histogram when all 30 samples agree, otherwise as the
per-sample vectors.

Repeat summaries retain one process summary per repeat and describe uncertainty
with the two-repeat process-level range. No confidence interval or claim of
statistical significance is manufactured from 30 samples inside one process.
Source, authored, and candidate identities must match between normal and
allocator processes for the same route/axis, workload, and repeat, and source
and authored identities must remain stable across repeats.

Route comparisons are deterministic versus memory-store/file-store on the same
workload, role, and repeat. Axis comparisons use the deterministic current
profile as the same-workload control. A change above 5 percent is a review
flag. These are contemporaneous route/profile comparisons and do not establish
a causal before/after speedup. Source-varying and authored-varying workloads
remain separate exact-case evidence; no cross-case average is emitted.

GNU `time` RSS includes the whole child process, setup, and report serialization.
It is retained as an observation and cannot establish a global bounded-memory
claim. The protocol remains claim-disabled (`performance_claim: none`).
