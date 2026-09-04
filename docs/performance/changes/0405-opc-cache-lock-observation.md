# Change 0405: OPC cache lock observation

Date: 2026-09-04

`performance_claim: none`

`claim_authorized: false`

## Problem and mechanism

The OPC cache diagnostics seam added by `a6fc94c99` exposes content-free
`DiagnosticSnapshot` pairs for the cache-state and same-Part flight-state
mutexes. The performance harness enables the `performance-diagnostics` crate
feature and adds the explicit `--opc-cache-lock-diagnostics` selector. The
selector is recorded as `configuration.opc_cache_lock_diagnostics`, while the
default worker path continues to call the source Part's unobserved `data()` entry
point.

The observer records cache and flight acquisition counts and elapsed
nanoseconds for each direct `Mutex::lock`. `Started` is sampled immediately
before the lock and `Finished` immediately after acquisition while the guard
is live. The caller-owned observer uses one nonallocating active sample and
rejects nested or unbalanced events. The OPC callback contract remains
synchronous, nonblocking, nonreentrant, and content-free.

The vectors cover every worker `PartData` request, including requests that
reach the cache before the coordinated source release. The benchmark's
`elapsed_ns` interval begins after that admission gate, so lock vectors are
descriptive per-sample evidence rather than intervals nested inside elapsed
latency. Condition-variable wait durations and mutex reacquisition are outside
the seam and are explicitly excluded. Package construction, prefill, source
I/O, gate coordination, and verification are also outside the direct-lock
scope.

## Report and comparator contract

Opt-in OPC contention rows add `source.opc_cache.lock_diagnostics` with exact
scope, exclusion, coverage, cache/flight vectors, and checked totals. The
parallel-metrics envelope derives `lock_wait_ns` only as the p50 of the
retained `total_lock_wait_ns` vector. Reports without this producer boundary
retain `lock_wait_ns: unavailable`; waiter counts are never converted into
time. The envelope remains `claim: "descriptive"` and is not a regression
metric.

`tools/perf_compare.py` validates the exact nested schema, unsigned integer
vectors, sample cardinality, cache-plus-flight totals, positive acquisition
counts for named contention rows, and the measured p50 cross-check. Instrumented
named contention rows must retain their OPC source envelope and one persistent
worker team. Unrelated rows in a combined report may omit the diagnostics.
Legacy reports that predate the selector normalize a missing configuration
field to `false`; `true` and `false` reports remain non-comparable.

## ADR compliance

| Requirement | Implementation boundary |
| --- | --- |
| [ADR 0003](../../adr/0003-snapshots-edits-and-patches.md) | The observer is read-path evidence only; cache, flight, pin, and publication semantics are unchanged. |
| [ADR 0005](../../adr/0005-io-memory-and-performance.md) | The scope names the exact direct lock boundary and keeps the normal path uninstrumented. No process-global thread count, waiter-to-time inference, or unsupported latency claim is added. |
| [ADR 0006](../../adr/0006-validation-security-and-compatibility.md) | Events contain no content, member names, credentials, or paths; malformed or incomplete diagnostic evidence fails closed in the comparator. |
| [ADR 0008](../../adr/0008-migration-and-verification.md) | Producer tests, schema tests, and an actual feature-enabled binary smoke are required before retaining evidence; this record does not certify a performance or compatibility result. |
| [ADR 0010](../../adr/0010-facade-archive-ownership.md) / [ADR 0011](../../adr/0011-ooxml-physical-package-ownership.md) | Observation is attached to the existing OPC cache owner and does not move archive or semantic ownership into the harness. |

## Remaining boundaries

The observer timer and callback dispatch add overhead to the opt-in run. Its
lock nanoseconds therefore cannot be compared with an uninstrumented run as a
pure contention or `Condvar` wait measure. A smoke report demonstrates schema
and scope only. Any later performance statement still requires clean release
ABBA measurements with the intended instrumentation configuration, retained
raw samples, stable CPU/environment identity, and the relevant allocation,
memory, and I/O evidence.

## Validation

The final release harness passes 256 tests with one ignored. The
[correctness manifest](../results/change-0405/correctness.json) binds the
commands, source hashes, and final logs; superseded failures remain labeled
as failures.

The focused Python comparator suite passes all 74 tests, including rejection
of boolean worker-team counts. The final full Python suite runs 850 tests
with 20 skips and no failures. The [release smoke evidence](../results/change-0405/validation.json) retains
24 normal and 24 instrumented rows across control/managed caches and one/two
workers, with three samples per row. Both reports pass the parallel-metrics
validator; normal rows omit lock diagnostics and observed rows retain them.
An unrelated selector rejects the flag. Source hashes, binary hash, toolchain,
and exact commands identify source commit `560e35035` and its freshly built
release binary. The tracked source is committed; untracked goal and pending
evidence files remain visible in the recorded dirty state. This smoke is
ineligible for a clean ABBA performance claim. Those measurements are evidence of instrumentation coverage and
schema validity only; no speedup or cache-latency claim is authorized by this
change.

## Harness baseline repairs

The full harness gate exposed stale golden output values. The DOCX source-edit
corpus pin predates the section XML hardening in `a260174e4`; two freshly built
corpora remain byte-identical, and the pinned SHA now identifies that current
writer output. The ODP and ODT batch publication expectations also predate part of the
sized-Deflate change in `3df216b46`. These are fixture maintenance issues and
cannot support before/after performance claims across differing corpus hashes.
The deterministic, semantic, member-preservation, and output-size assertions
remain required.

The XLSX vendor-extension negative control now inserts a compact in-document
comment into `styles.xml`. Its old trailing-space mutation was rejected by
the authored XML policy before the untouched-member oracle could inspect it.
The corrected probe still changes the supposedly untouched member and must
therefore fail that preservation oracle.

The row-visibility refusal probe now requests a real hide/unhide before
checking refusal, because edit-handle creation is lazy. Production commit
`01b42f487` restores the row-visibility formula guard in its existing
namespace-aware scan; all eight focused row-visibility tests pass. The probe
requires an error,
an empty output sink, and unchanged source identity and bytes. The cell-value
managed context records the source payload ceiling separately from a fixed,
checked 64 KiB publication-planning allowance. Its prior exact payload limit
was exhausted by a 3,048-byte physical-member lookup reservation. The same
explicit allowance is used by the untimed output-refusal replay, which still
verifies the actual first write size, one-under typed output limit, zero output
charge, and source identity. The configuration identity records the allowance
so payload-only historical reports cannot be silently compared with this
context.
