# Change 0133: filesystem operation-metrics envelope

Date: 2026-08-15

Status: schema and correctness evidence only; no performance claim

## Scope

The existing isolated filesystem child selectors now add an
`operation_metrics` object to each warm or `cold-requested` `CaseResult`.
Schema version remains `1`; non-filesystem results omit the additive object,
and the raw top-level `filesystem_evidence` sample records are preserved.

The envelope is built only from already-collected child evidence. It aligns
each vector with the sorted `elapsed_ns.samples` vector and records an explicit
`sample_count` and `alignment` (`elapsed_ns.samples`). Source vectors include
logical read calls, requested bytes, returned bytes, and maximum concurrent
reads. Process vectors include user/system CPU ticks, the procfs clock factor,
faults, context switches, RSS deltas, and after-sample high-water RSS.

## Interpretation boundary

Procfs CPU, fault, context-switch, and RSS values are operation deltas between
the child’s before/after snapshots. `peak_rss_bytes` is the process-lifetime
high-water `VmHWM` observed after the operation; it is not an operation peak.
`rss_delta_bytes` is also not a peak. The source and optional publication,
output-length, and materialization vectors fail closed when their per-sample
option cardinality is asymmetric. A measured zero remains a numeric zero.
Unsupported or unavailable metrics omit numeric vectors and expose an explicit
`status`; they are never encoded as zero or JSON `null`.

No allocation, copied-byte, decompressed-byte, recompressed-byte, or physical
I/O claim is introduced. Those quantities are not instrumented by this
harness. The change does not alter production crates, timed operations, the
default 36 cases / 198 records, or comparator policy.

## Verification

Focused unit tests cover warm/cold partitioning and elapsed ordering, exact
cardinality, measured zero versus unavailable, asymmetric optional fields,
not-applicable scopes (including eager OPC's uninstrumented `fs::read` path),
and status/max-value serialization. A one-sample
release filesystem smoke covered eager/source OPC and CFB saves plus eager/
source PPTX and DOCX opens. The smoke confirmed schema version `1`, exact
sample cardinality, measured source/process vectors, and omitted numeric
vectors for not-applicable source scopes.
