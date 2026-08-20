# Change 0230: operation write and procfs I/O metrics

Date: 2026-08-20

Status: schema and correctness evidence only; no performance claim

## Scope

The standalone `tools/perf-baseline` harness now promotes the existing logical
sink summary into the additive `operation_metrics.sink` envelope for cases that
already expose a deterministic top-level `sink` summary. The promoted vectors
are `accepted_bytes`, `write_calls`, `largest_write`, and the six fixed
`write_size_buckets` vectors. They are aligned with `elapsed_ns.samples` and do
not alter the timed workload or remove the existing top-level summary. Schema
version remains `1`.

The filesystem operation envelope now exposes all existing `/proc/self/io`
operation deltas alongside its CPU, fault, context-switch, and RSS vectors:
`rchar`, `wchar`, `read_bytes`, `write_bytes`, `cancelled_write_bytes`, `syscr`,
and `syscw`. Their vector scope is
`child_process_interval_delta_including_procfs_probe_overhead`: the after-
snapshot procfs read can itself add `rchar` and `syscr`. Procfs collection
remains best-effort; unavailable values retain their explicit status and omit
numeric vectors.

## Interpretation boundary

Sink metrics describe logical lengths accepted at the harness
`Write::write` boundary. They do not describe requested lengths, rejected
calls, operating-system syscalls, storage I/O, memory copies, or writer-
internal buffering. The harness does not retain requested-versus-accepted
pairs for short writes, so no requested field is synthesized.
`operation_metrics.sink.output_bytes` remains the separate final output-length
observation and is never inferred from accepted sink bytes; seekable sinks may
accept rewrites.

Only a deterministic per-case summary is promoted. The summary is repeated
over the aligned elapsed samples because the measured sink summaries were
already checked equal across those samples; this is not a new per-sample
measurement. Filesystem selectors have no logical sink summary and retain
`not_applicable` sink-write vectors. No allocation, decompressed-byte,
recompressed-byte, operation-only procfs, or physical-I/O claim is introduced.

## Verification

Focused operation-metrics tests cover all seven procfs I/O fields, measured
zero versus unavailable serialization, accepted-boundary sink vectors and
buckets, alignment across multiple samples, omission of requested-length
placeholders, and rejection of an empty sink sample vector. Non-building
verification for this isolated batch is `git diff --check`; the coordinator
will run the repository build and test gates separately.
