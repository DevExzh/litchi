# Change 0240: explicit parallelism metric envelope

Date: 2026-08-20

The standalone performance harness now emits a top-level `parallel_metrics`
envelope. It records the configured worker-budget selection and, for explicit
scaling results, the configured worker width plus the deterministic logical
task count already exposed by the harness. OPC cache contention results can
also report the width of one harness-created local worker team when its team
creation count is exactly one.

The envelope is intentionally fail-closed. It does not inspect process-global
thread lists, read `/proc`, infer workers from CPU utilization, or treat
`waiter_joins` as lock time. Range-simulation results expose an exact
per-sample physical request/chunk vector; results without that explicit vector
leave deterministic chunk count unavailable with an explicit scope/reason.
Process thread count and lock wait remain unavailable. The range, CFB
selective, and CFB `open_stream` simulators expose their exact per-sample
physical request counts; other source requests and byte-size buckets are not
converted to task or chunk counts. The local worker observation is not a claim
about all threads in the benchmark process.

Focused module tests cover valid scaling and worker-team records, malformed
worker budgets, and the absence of lock/chunk inference. No Cargo command was
needed for this instrumentation review; formatting is checked with rustfmt.
