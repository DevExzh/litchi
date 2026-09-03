# Change 0385: OPC source materialization benchmark boundary

Date: 2026-09-03

Status: implemented

`performance_claim: none`

`claim_authorized: false`

## Scope

The standalone performance harness now exposes the opt-in
`opc_source_materialize` selector. Deterministic source-backed OPC catalog
construction and validation happen before timing. The timed operation contains
only conversion into a complete owning `OpcPackage`; full Part, relationship,
content-type, and payload-digest verification happens after timing.

Each retained sample reports the logical source calls and returned bytes made
by conversion, the materialized Part count, operation-scoped process counters,
and allocator counters when the allocator-enabled binary is used. The normal
binary reports allocator data as explicitly unavailable rather than as a
measured zero. Sink counters are explicitly not applicable. Requested range
sizes, copied bytes, compressed/decompressed/recompressed bytes, physical I/O,
and operation-local peak RSS remain unavailable where the instrumentation
cannot observe them; none are inferred. A dedicated
`evidence_only_opc_source_materialization` latency identity and paired
`in_process_instrumented_source_read_at` scope keep these rows outside latency
comparison while allowing measured operation vectors to be validated.

The selector is excluded from `Case::DEFAULT`, so the default matrix remains
36 selectors and 198 result rows. Its corpus construction, source-backed open,
verification, and result serialization remain outside the measured interval.

## Verification

The 24-test operation-metrics suite includes multi-sample tied-elapsed ordering,
and the focused selector test exercised both tiny compressible and few-large
incompressible corpora. Together they verify opt-in status, source and
materialization vectors, explicit sink status, normal-binary allocator status,
and stable sample alignment. The comparator and ABBA suites passed 136 tests,
including dedicated claim/scope acceptance and mismatched/comparable-claim
refusal. One actual five-sample Rust report passed the Python operation-envelope
validator.

Independent release smoke runs exercised the normal and allocator-enabled
binaries on both shapes. The allocator binary reported measured allocation
vectors; the normal binary reported them unavailable. A release default-matrix
smoke still produced 198 rows and no `opc_source_materialize` row. The existing
source-backed OPC suite passed 128 tests, both release binary checks passed,
and `git diff --check` passed.

Validation used the explicitly selected installed stable toolchain because the
repository-pinned 1.95 toolchain lacks Cargo in this environment. This change
adds a truthful measurement boundary; it does not itself establish a before/
after performance result.
