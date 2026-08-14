# Change 0108: process CPU and context-switch evidence

Date: 2026-08-14

Status: correctness and schema evidence only; no performance claim

## Scope

The standalone `tools/perf-baseline` filesystem child now records additional
Linux procfs counters in its existing optional `process_metrics` object:

| Field | Source and unit |
|---|---|
| `user_cpu_ticks` | `/proc/self/stat` field 14, process user-mode CPU clock ticks |
| `system_cpu_ticks` | `/proc/self/stat` field 15, process kernel-mode CPU clock ticks |
| `clock_ticks_per_second` | `rustix::param::clock_ticks_per_second()`, ticks per second |
| `voluntary_context_switches` | `/proc/self/status`, operation delta |
| `nonvoluntary_context_switches` | `/proc/self/status`, operation delta |

The CPU and context-switch values are component-wise saturating deltas between
the existing before/after process samples. `clock_ticks_per_second` is retained
from the after-sample because it is a unit conversion factor rather than an
operation counter. Parsing and arithmetic reject malformed or overflowing
procfs values; the existing caller still converts an unsupported or unavailable
procfs sample to `None`.

## Interpretation boundary

Process CPU utilization, when derived by a consumer, is

```text
(user_cpu_ticks + system_cpu_ticks) / clock_ticks_per_second / elapsed_seconds
```

This is process CPU time normalized by elapsed wall time. It is not
automatically whole-machine utilization, CPU busy percentage, or scaling
evidence. No result values or speedup claims are included in this change.

## Verification

The focused module tests cover a command name containing closing parentheses,
status lines with parenthesized text, saturating CPU/context deltas, and
explicit serialized units. The existing process-metrics fallback remains
best-effort through `Snapshot::read().ok()` at the filesystem child caller.
