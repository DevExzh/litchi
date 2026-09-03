# Change 0386: full-run resource profile artifact

Date: 2026-09-03

Status: implemented

`performance_claim: none`

`claim_authorized: false`

## Scope

The scheduled and manually dispatched full performance job now runs the
existing process-isolated resource profiler against the release harness built
by the immediately preceding workflow step. It records the explicit non-iWork
workload set for OPC source publication, managed XLSX batch publication, RTF
streaming creation, CFB selective read and save, and OPC/CFB 1/2/4/8/available
worker scaling. The current-head run uses one warmup and three retained harness
samples; external tools remain one-sample, process-total diagnostics with their
existing overhead caveats.

The wrapper probes GNU time, `perf`, `strace`, Heaptrack, and `taskset` but does
not install them. Missing, permission-denied, unparsed, and measured states are
retained explicitly rather than converted to zero. The step writes
`target/perf/resource-profile.json`, validates its schema/tool identity,
non-iWork scope, exact workload list, sampling configuration, successful
harness reports, logical measurements, tool descriptors, and supported or
unsupported external-profile status shapes, then uploads it in the full job's
always-run artifact.

The command deliberately reuses the prebuilt binary and does not invoke the
wrapper's `--build` path. The report therefore retains the wrapper's
conservative prebuilt-binary provenance classification and exact binary hash;
workflow step ordering is not presented as a cryptographic source-to-binary
proof.

## Verification

The workflow-policy and resource-profiler suites passed 70 tests. Policy
coverage pins the full-job-only command, release-binary path, explicit workload
set, sampling and timeout bounds, output validation, graceful unavailable-tool
handling, successful nested Heaptrack parsing, always-run upload, and push/PR
path triggers. The workflow parsed successfully as YAML and `git diff --check`
passed.

This change supplies current-head diagnostic artifacts and scaling summaries;
it establishes no before/after speedup, cold-cache, remote-I/O, operation-local
RSS, decompression, recompression, or copied-byte claim.
