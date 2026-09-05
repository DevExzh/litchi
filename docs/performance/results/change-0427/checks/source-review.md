# Change 0427 source review

This is an independent source-only review of the allocator retention lane and
its report verifier. I read the 0427 acceptance checklist and retention design
note, with ADR 0003, ADR 0005, and ADR 0006. I did not run Cargo, tests,
profilers, capture commands, or CPU workloads.

The lifecycle mechanics are sound in the current frozen Rust draft. The region starts
before the baseline and ends after the sink drop; checkpoint reads do not
allocate; output is compared with the prevalidated role-specific bytes before
any drop; reallocations are included in the checked allocated-minus-deallocated
balance; owned snapshots are dropped before their packages; and the
source-backed editor is consumed by publication, with the plan and source view
released before the caller source owners. Invalid, overflowed, unavailable, or
unbalanced observations fail the capture instead of being serialized as zero.

## Findings

### Resolved: report schema and provenance

The frozen Rust `Report` now includes `corpus_manifest`, allocator identity,
and instrumentation identity, and the verifier’s exact top-level schema and
manifest checks match those fields. The report carries source, destination,
expected-output, binary, and observer identities; the capture/replay custody
bundle supplies the external protocol, source-manifest, and verifier hashes.

### Resolved: verifier owner label includes live temporary ReadAt Arcs

At `pptx_retention.rs:702-711`, `source_read` and `destination_read` are
cloned before `inputs_and_sink_ready` is captured at line 708. Those
`Arc<dyn ReadAt>` values remain live through that checkpoint and are consumed
only by the following `from_read_at` calls. The Rust phase description at
lines 199-201 and `verify-report.py:81` now both name the temporary ReadAt
Arcs, so the verifier accepts the actual source-backed ownership boundary.

### Resolved: source-backed drop order and phase text

The frozen producer and verifier now agree that publication consumes the editor,
then the plan and source-backed view are released, followed by the caller
`InstrumentedSource` Arcs. Owned snapshots are still dropped before their
packages. The source-backed Arc count remains checked at exactly one owner
after document-handle drop.

### Resolved: absolute counter-balance verification

`verify-report.py:283-300` now validates the absolute allocated-minus-
deallocated balance and lifetime high-water relation for every checkpoint,
including the baseline, before checking monotonic transitions.

### Resolved: repeated output identity in replay

`docs/performance/results/change-0427/verify.py:82-83` now binds expected
output hash and byte length across repeats of each `(api, corpus)`, in addition
to cross-API source/destination identity. Root is retaining the adversarial
probe work separately; the source lane itself compares every emitted sink with
the prevalidated bytes before any drop.

## Review result

The allocator boundary and explicit drop sequence are suitable for the stated
descriptive callback-order observation with the producer and verifier phase
labels aligned. No performance, RSS, cache,
managed-budget, leak, or object-owned-memory claim is supported by this lane.
