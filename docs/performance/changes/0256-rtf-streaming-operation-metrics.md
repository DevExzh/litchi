# Change 0256: operation-scoped RTF streaming resource metrics

## Status

Landed in `83326fbdc`. The harness can now collect descriptive resource
observations for fixed-window RTF creation. No new resource result or
optimization claim is accepted by this change alone.

## Measurement boundary

Each retained `rtf_streaming_create` sample now has one aligned
`operation_metrics` observation:

- procfs-before sampling occurs before the timer;
- the allocation region begins immediately before `Instant::now()`;
- the elapsed interval contains only `write_streaming_rtf`;
- duration is captured before allocator finish and procfs-after sampling;
- sink digest finalization and post-operation harness counter/correctness
  gates remain outside the elapsed interval; the writer's own invariant checks
  and `finish()` remain inside the timed call.

Observations are ordered by `(elapsed_ns, sample_index)`, matching the harness
statistics ordering and retaining an explicit identity for timing ties. Mixed
provider presence, mixed allocator status, missing measured fields, and
overflow asymmetry fail closed.

## Metric semantics

- The allocator binary publishes checked allocation/deallocation/reallocation
  calls and bytes plus absolute live/high-water values before and after the
  operation. The ordinary binary leaves the optional allocation envelope
  absent.
- Process fields are best-effort same-process procfs deltas. Their scope
  explicitly includes after-snapshot probe overhead; unsupported platforms
  report `unavailable` without fabricated zero vectors.
- `peak_rss_bytes` remains the process-lifetime high-water mark, not an
  operation-local peak. Allocator live and peak values are also absolute
  process counters rather than isolated-operation peaks.
- Source, materialization, publication, and CFB phase fields are explicitly
  not applicable to fresh in-process sink creation. Existing deterministic
  logical sink-write vectors remain measured.

## Verification

Focused operation-metrics tests passed for measured and unavailable providers,
out-of-order alignment, allocator vectors, sink replication, and JSON
serialization. The fixed-window RTF benchmark test passed with resource
metric presence/alignment checks. Rustfmt and `git diff --check` passed, and an
independent read-only review found no remaining timing or schema blocker.

## Remaining evidence

Run the allocator target in a pinned release environment for the deterministic
tiny/medium/large RTF shapes with the normal warmup/sample protocol, retain the
machine-readable report and executable identity, and audit allocation/RSS
vectors before making a bounded-memory or allocation-reduction claim. The
existing change 0097 latency result remains scoped to ASCII batching and does
not acquire a resource claim retroactively.
