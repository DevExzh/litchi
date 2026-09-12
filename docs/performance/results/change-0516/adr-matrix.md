# 0516 ADR compliance

All accepted ADRs and the index remain bound by `adr-manifest.json`. This
matrix describes the archived `after` candidate; its retention decision is
separate from semantic compatibility.

| Constraint | Candidate behavior | Evidence |
| --- | --- | --- |
| 0001, 0004: typed API layers | Reuses a private worksheet parser; no public API or raw archive types are introduced. | Candidate patch and owner checks. |
| 0002, 0024: dependency ownership | All production changes remain within `litchi-xlsx`; the separate probes use public APIs. | Source manifests and crate-boundary checks. |
| 0003: immutable snapshots and atomic edits | Feeds only effective changed outputs requiring Store verification. Exact no-ops and source-layout planning retain their existing paths; inverse, source conflict, and publication checks remain. | Transaction patch, differential edit tests, all-features suite. |
| 0005: resource use and evidence | Discards provisional state on admission failure and parses exact output through the existing owner. The 128 MiB accounting threshold covers output capacity plus new parser/proof state, not process RSS or baseline compactor allocations. | Resource reviews, repeated allocator lanes, native RSS logs. |
| 0006: preservation and validation | Feeds emitted normalized events, retains exact compacted bytes, defers provisional errors, and preserves grid, web, style, requested-change, and publication order. | Final implementation review; namespace, malformed-input, extension, and error-order tests. |
| 0011: OPC ownership | Package ownership and final publication stay with the existing OPC path. The independent fixture repair accounts for its existing 64 KiB exact-copy buffer. | Baseline test review and unchanged-production full suite. |
| OOXML format ownership | MCE/x14ac, shared-string dependencies, and shared-formula expansion use the authoritative exact-output fallback. The 4,096-cell/1 MiB Store handoff remains unchanged. | Input/event/output gates, fallback probe, callback and extension tests. |

No ADR, unsafe-code policy, executor policy, or format support contract changes.
OLE2 and OOXML retain priority until their optimization goal is complete; ODF
is deferred and iWork is outside this workstream.
