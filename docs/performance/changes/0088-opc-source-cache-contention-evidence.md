# Change 0088: OPC source-cache Budget and contention evidence

Date: 2026-08-13

Production revision: `d488ed128`

Status: opt-in harness and release ABBA structural/distribution evidence complete;
no performance improvement accepted

## Scope

The performance harness adds three opt-in selectors over one fixed 256-Part,
1,024-byte-per-Part, incompressible OPC corpus:

- `opc_source_cache_budget_boundary` emits exact-budget success and
  one-byte-under managed refusal records;
- `opc_source_cache_control_contention` exercises the compatibility cache with
  finite `SourceCacheLimits`; and
- `opc_source_cache_managed_contention` exercises the same cells with an
  explicit `ExecutionContext` and hierarchical memory `Budget`.

The contention selectors cover same-Part and fixed-work disjoint-Part waves at
`1/2x`, `1x`, and `2x` capacity. Worker widths are the existing capped,
deduplicated `1,2,4,8,available` selection. With five resolved widths, the
matrix contains 62 records: two managed boundary records plus 30 control and
30 managed contention records.

Every cell creates one worker team and reuses it across all warm-ups and
samples. Each iteration uses a fresh source and package, fills a disjoint cache
working set, admits the initial timed cohort through an explicit source gate,
and starts timing only immediately before gate release. Each newly encountered
timed compressed payload range receives a fixed 10 ms delay once. Package open,
prefill, worker construction, rendezvous and semantic verification stay outside
timing.

## Required invariants

The harness fails rather than emitting a record unless all applicable
invariants hold:

- the exact managed Budget admits one Part and retains its reservation through
  the cache after the returned handle is dropped;
- a one-byte-under Budget reports exactly two reservation failures, retains
  nothing and performs zero payload I/O;
- same-Part waves expose one flight, `workers - 1` waiters, one cold read and
  one shared allocation;
- disjoint-Part waves expose one initial flight and one simultaneous fixed
  source delay per worker (`worker_count` flights and delays), then load every
  Part in the fixed working set once;
- `1/2x`, `1x`, and `2x` cells report their exact eviction, bypass, retained
  entry and retained byte counts while all returned handles remain pinned;
- compatibility diagnostics remain unmanaged, while managed diagnostics and
  the caller-visible Budget agree exactly;
- dropping returned handles leaves only retained cache reservations, and
  dropping the package releases all remaining reservations; and
- one and only one persistent worker team is created per contention cell.

The JSON report keeps per-sample cache counters, occupancy, Budget diagnostics,
gate arrivals/concurrency, pre-release flights/waiters, and post-drop Budget
use. Same-Part widths change the number of requests, so they report request
throughput only and mark Amdahl analysis not applicable. Disjoint widths retain
one fixed request count and report speedup, efficiency and an Amdahl serial
fraction with explicit baseline, valid, superlinear, slowdown, or out-of-model
classification.

## Claim boundary and acceptance gate

The coordinated 10 ms source delay is a deterministic test instrument. It does
not model production storage, scheduler, allocator or decompression latency.
No performance improvement, regression, scalability limit, memory saving or
control-versus-managed delta is claimed from this harness implementation or a
debug smoke.

Any performance conclusion requires a clean release build, fixed source and
compiler revisions, CPU affinity, balanced control/managed ABBA ordering,
retained raw distributions, stable counter identities, allocation counts,
peak heap and RSS evidence, copied/decompressed-byte and CPU-utilization
counters, contention profiles/flame graphs, and disclosure of variance and all
out-of-model Amdahl cells. The production cache implementation is unchanged by
this tranche.

## Release evidence

The follow-up release capture is retained under
[`results/opc-source-cache-release-abba-0100/`](../results/opc-source-cache-release-abba-0100/).
It was built with `--release --locked` from a clean archive of committed
revision `a1b692297b2493a3f523aa064e6be366271c4f52`, pinned to CPUs `0-11`,
and ran the balanced order `control-A`, `managed-A`, `managed-B`, `control-B`.
Each of the 30 cells per leg used 3 warmups and 30 measured samples at resolved
widths `1,2,4,8,12`. The exact-budget and one-byte-under boundary records also
passed, including zero payload I/O for the refused one-byte-under case.

Independent recomputation matched the retained hashes, percentiles, means,
Student-t intervals, all 60 directional Welch intervals, and every cache,
source-I/O, flight, waiter, pin and Budget invariant. Zero cells passed both
directional confidence gates, so no managed-versus-control speedup is accepted.
The 10 ms source delay remains a deterministic coordination instrument rather
than production latency. Per-sample allocation counts, peak RSS and attributable
hardware counters were not captured and remain explicit evidence gaps. The
reviewed 34 MiB executable and temporary build tree are intentionally excluded
from version control; their provenance hashes remain in
`results/raw-manifest.json`.

This tranche also does not close ADR 0005's broader execution-context adoption
work. Compatibility source-backed constructors remain deliberately unmanaged,
and the managed lazy Part path reserves declared payload memory but does not
charge the context's `InputBytes`, `Work`, or `Objects` dimensions for source
reads and decompression; separate finite `ReadLimits` still bound that work.
Format-level parsed-cache adoption likewise remains outside this harness-only
change.
