# Change 0309: OPC source-backed path ingress

Status: implemented; focused validation complete

`performance_claim: none`

## Path-backed source ingress

`SourceBackedPackage::from_path` and
`SourceBackedPackage::from_path_with_limits` open a filesystem path through a
positional `FileSource`. The source owner captures the package identity and
revision, validates the OPC catalog and relationship graph, and retains the
source for later bounded positional reads.

Opening a source-backed package is catalog-only for ordinary package payloads:
ordinary part bytes are not materialized merely because the path was opened.
Individual part reads remain source-bound and are checked against the captured
source state before data is returned.

## Finite resource policy

The no-argument path constructor uses the finite default read and deferred-payload
cache policies. Callers handling untrusted or large inputs can select
`from_path_with_limits` to provide an explicit bounded read policy; cache
limits remain an independent policy where the source-backed package exposes
them.

## Eager CRUD boundary

`SourceBackedPackage` is the immutable, source-bound ingress and deferred-read
owner. Eager `OpcPackage` remains the explicit boundary for ordinary mutable
CRUD, whole-package materialization, and publication. Crossing that boundary
is deliberate and subject to the caller's read/materialization limits; opening
a path through the source-backed constructor does not silently construct an
eager mutable package.

## Regression scope

Validation covered 242 OPC library tests and 3 focused filesystem path tests.
Strict OPC library-and-test Clippy and rustdoc checks also passed. Final
repository handoff still expects rustfmt and diff-hygiene checks. This record
does not claim a complete filesystem matrix, publication round-trip coverage,
or coverage for every OPC part family.

## Measurement boundary

No total-RSS, peak-memory, or throughput improvement is claimed. This change
also does not alter ordinary eager `OpcPackage` constructors or claim that
their materialization behavior has changed.
