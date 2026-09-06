# 0445: Matched plain-source OPC Part-addition baseline

Add opt-in `opc_part_add_plain_lifecycle` beside the observed lifecycle. One run
function prepares either source before timing, then executes identical catalog
opening, topology planning and consuming sequential publication to a hashing
discard sink. Plain source counters are explicitly unavailable, with no values.
The same fixture, output digest, sink summary and correctness gates bind both.

The single-build observed/plain ABBA matrix has 24 reports/720 retained samples,
CPU 2, one worker, normal and allocator modes, 64/1024/4096 Parts, and 30 samples
with three warmups. Plain normal p50 is 1.074–1.079 / 5.230–5.280 /
18.869–18.886 ms. Medium and large are respectively 34.970–35.524% and
69.576–69.635% below observed mode. This measures observer overhead and removes
no production work. Allocation calls, requested bytes and above-entry peaks are
identical. One observed-allocator tiny p99 repeat flag (-7.928%) remains visible.

Four profiles show the observed reader at 68.989% of run-frame self period.
Plain content-type map parsing covers 29.291% inclusive of the run-frame subset.
That subset includes setup/probes/warmups; inclusive rows overlap. Source review
identifies an owned Part-name String cloned into PackURI as the next allocation
candidate. Do not skip source/candidate validation or add an uncharged manifest
cache. No production, bounded-memory, native, cold/range or scaling gain is claimed.

Verification: 373 harness tests and 37 existing topology tests passed (one
existing harness test ignored); workspace/harness checks, warning-denied rustdoc,
formatting and boundaries passed. Strict harness lint retains only inherited
debt. All six ZIPs match 0444 byte for byte; 12 pilots and 53 independent corruption
probes pass. A profile-row tie-ordering replay failure was corrected and retained;
raw measurements were unchanged. The final protocol froze before formal capture.

Only three standalone harness Rust files change. The registry is 438 selectors,
with 36 defaults and unchanged semantic representative coverage. The full non-iWork
goal remains active. See [measurements](../results/change-0445/measurements.md),
[hotspot review](../results/change-0445/hotspot-review.md), and
[evidence/replay instructions](../results/change-0445/README.md).
