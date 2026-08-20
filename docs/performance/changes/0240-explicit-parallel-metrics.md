# Change 0240: explicit parallelism metric envelope

Date: 2026-08-20

The standalone performance harness now emits a top-level `parallel_metrics`
envelope marked `claim: "descriptive"`. It records the configured
worker-budget selection and, for explicit
scaling results, the configured worker width plus the deterministic logical
task count already exposed by the harness. OPC cache contention results can
also report the width of one harness-created local worker team when its team
creation count is exactly one.

The envelope is intentionally fail-closed. It does not inspect process-global
thread lists, read `/proc`, infer workers from CPU utilization, or treat
`waiter_joins` as lock time. Range-simulation results expose an exact
per-sample physical request/chunk vector, aligned through the serialized
`elapsed_ns.sample_order`; results without that explicit vector leave
deterministic chunk count unavailable with an explicit scope/reason. Worker
widths are rejected when they are absent from the configured worker budget.
The CFB selective `read` phase and CFB `open_stream` per-operation sums exclude
their timed-open requests. Process thread count and lock wait remain
unavailable. Other source requests and byte-size buckets are not converted to
task or chunk counts. The local worker observation is not a claim about all
threads in the benchmark process.

Focused Rust-module and Python-comparator tests cover valid scaling and
worker-team records, malformed worker budgets, sample alignment, metadata
shape, result cross-checks, and the absence of lock/chunk inference. The
comparator, ABBA summary, and ABBA package validator reject malformed
`parallel_metrics` when a schema-v1 report emits it; none compares this
descriptive envelope as a speedup or regression metric. No Cargo command was
needed for this instrumentation review; formatting is checked with rustfmt.
