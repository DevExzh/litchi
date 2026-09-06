# 0445: Matched plain-source OPC Part-addition lifecycle

Previous turn: progress. Change 0444 committed an observed Part-addition baseline,
but 55.24% of whole-process self samples landed in the instrumented source reader.
Its ordinary-range accounting scans every member range on every read. Before
attributing the 4096-Part time curve to production topology, separate that observer
from the lifecycle. This is a measurement enabler, not a production optimization.

Add opt-in `opc_part_add_plain_lifecycle` beside the existing observed selector.
Both use the same fixture/output and gates, source-backed catalog open, topology
plan construction, consuming sequential publication and hashing-discard sink.
The mode branch, OwnedSource/input clone or observer setup, and source counters
are outside timing. The timed body is shared. Plain reads are explicitly
unavailable, not fabricated zeroes; sink, allocation and process observations
retain their actual statuses. Both selectors remain outside the default set.

Capture a single current build in an observed/plain ABBA order: A1 observed R1,
B1 plain R1, B2 plain R2, A2 observed R2. Each phase has normal and allocator modes,
three sizes (64/1024/4096 Parts), 30 samples/three warmups, CPU 2 and one worker:
24 reports/720 samples. R1 is normal then allocator, tiny to large; R2 reverses
mode and shape order. Describe observer overhead separately from production
performance. Review all absolute 5% repeat flags and every paired result.

Retain independent archive/report verification and mutation probes. Exported ZIP
hashes must match 0444 exactly. Reuse its fixture formula, typed Rust gates, raw
untouched records and exact manifest/relationship checks. Freeze the final oracle,
capture method and actual timestamp only after code/fixture/pilot validation and
before formal samples. Retain one plain and one observed whole-process stat and
record profile (four profiles); use symbol-level stacks without addr2line inline
expansion. Analyze plain stacks to select the next production bottleneck; do not
claim operation-local or native/cold/range/scaling attribution from these profiles.

No production Rust/API/dependency/unsafe/scheduler/ambient I/O changes. Accepted ADR
tree c950b6c8be822561b498d7bbe87c460873dcbf49 is unchanged from the prior complete read.
Root CPU jobs stay serialized and Rust stays fixed while a job is live. The full
non-iWork goal remains active; semantic ownership, repackaging, native breadth,
cold/range sources and real scaling remain open.
