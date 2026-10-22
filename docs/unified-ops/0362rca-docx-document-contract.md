# 0362rca — DOCX output memo ownership repair

This follow-up repairs the document contract's output-fingerprint ownership.
The eager owner now keeps its OPC graph and complete artifact digest in an
OPC-specific value wrapper. The weak graph key and digest are inline in each
package clone: cloning copies the memo for that immutable graph, while
`Arc::make_mut` dissociates the live weak key before any in-place mutation.
Core and custom properties are reloaded from the candidate OPC graph at every
package edit, and `write_to` serializes that graph, so the OPC state remains
the complete output dependency for metadata as well as document parts. No
generic `Shared<T>` field allocates an invalidation counter or atomic.

The source owner now uses the memo already retained by an owned
`SourceArtifact`. A version-keyed owner memo is allocated only for generic or
managed physical sources. Effective prepared topologies receive a separate
once-only output memo, after checking source freshness; their composition
fingerprint therefore does not re-emit and hash the same topology on every
compose/apply call. The old unconditional per-owner source mutex and the
unconditional effective write-and-hash path are gone.

The raw Tree, package composition, owner-proof scope, reopened equal-content
replay, changed-untouched-part refusal, and source-freshness checks from
0362rc remain unchanged. The focused raw reopen/forward/inverse/second-edit
coverage remains in the source tree. First, repeated, and alternating eager
package lifecycle measurements, plus the effective-topology paired batch
probe, are pending the coordinator's released build lane; this commit makes
no unmeasured performance claim.

Validation in this constrained lane was limited to formatting and static
source review. No Cargo build was started while the coordinator's target and
disk budget were reserved by other lanes.
