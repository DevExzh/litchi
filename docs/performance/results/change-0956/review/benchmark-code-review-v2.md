# 0956 benchmark code review v2

Reviewed the finalized benchmark source snapshot on 2026-10-07.
Source SHA-256:
eff67c2e67e8fd634ee689f67a6405f365cf707e5cc977b350ce3d43930fca28

No checkout edits, builds, tests, or workload runs were performed.

## Blocking source issue

The current snapshot does not compile: CasePlan is declared at
tools/perf-baseline/src/xlsx_managed_compression_benchmark.rs:209-230 without
a preflight_actual_commit_inverse_verified field, but build_case_plan initializes
that field at line 634 and run_case reads it into CaseReport at line 675.
Add the field to CasePlan before accepting the integration build. This is the
only material blocker found in this final snapshot.

## Resolved review points

A fresh BoundedSink is now created for every warmup and measured iteration
(lines 638-652), so capacity is not reused across samples. Its configured
output cap is checked independently of Vec capacity (lines 120-175).

Output ownership is now acquired atomically with create_dir (lines 717-739),
and failure.json is written only when this invocation created the directory.
Duplicate warmup/sample flags are rejected. The report explicitly states that
managed context and source cloning are setup outside the clock, and the
precomputed per-sample inverse is labeled fixture-only. The untimed parent
qualification now records and reports an actual managed-commit inverse proof.

The earlier editor-teardown concern does not apply: publish_multi_commit_to_stream
takes self and consumes the editor/package before returning
(crates/litchi-xlsx/src/cell_values/source.rs:599-604). The measured publication
therefore includes that consumed editor/package teardown; commit destruction is
a separate commit_drop_ns phase and is included in total_ns.

The source hash is intentionally documented as immutable fixture identity.
The timed source is an OwnedSource clone of that fixture, and no freshness
mutation claim is made. The timed inverse uses a precomputed fixture patch,
while the untimed parent preflight verifies the actual managed commit inverse.
