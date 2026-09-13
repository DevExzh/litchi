# XLSX sorted provenance merge: compatibility boundary

Status: candidate preparation; no production adoption or performance claim.

The thirty ADR/index files are byte-identical to the files read earlier in this
session. `preparation.json` binds the checked 0555 manifest and current source.
The optimization priority remains OLE2 and OOXML; ODF is deferred and iWork is
outside this coordinator's scope.

| Constraint | Candidate boundary and required evidence |
| --- | --- |
| ADR 0001/0002/0024: priorities and ownership | Private XLSX Store construction only; no facade, package, or dependency change. |
| ADR 0003: immutable snapshots and atomic publication | Move owned parsed cells and clone selected source cells; no source mutation, shared mutable state, or publication change. |
| ADR 0005: resources and measured performance | Preserve checked combined count and fallible reserves. No new persistent state or runtime. Fresh matched native, allocation, and instruction evidence is still required before admission. |
| ADR 0006: preservation and validation | Complete rewritten-output validation and reduced parse remain authoritative. Preserve every Stored field, worksheet structure, extents, merge behavior, and complete-parse fallback. Differential Store tests supplement existing public preservation tests. |
| ADR 0008: verification state | Preparation tests are separate from an adoption decision. Keep the production source at its baseline until measured gates pass. |
| ADR 0010/0011: physical package ownership | No ZIP, OPC, physical identity, compressed-member, or save-path changes. |
| Safety and concurrency | Safe Rust only, no new threads, allocator, provider, cache, or global runtime. No new scaling or concurrency claim. |

The source proof and candidate review must verify constructor ordering, omitted
rectangle semantics, duplicate/refusal behavior, and all index/extent inputs.
Allocation-failure injection, fuzzing, and native-producer capture are not
implied by ordinary unit and integration test success. Their availability and
applicability must be recorded at measurement/adoption time.
