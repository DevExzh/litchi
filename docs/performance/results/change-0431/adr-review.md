# Change 0431 ADR review

This review covers the compressed source-part token, its OPC authorization,
PPTX publication integration, and the matched evidence protocol. The refined
comparison is a scoped accepted API-duration optimization for the named
source-backed workflow: 16 processes and 480 samples pass per role with no
API repeat or regression flag above 5%. Portable verification passed before
and after task scratch cleanup, which also passed.
No global-goal or whole-process memory claim is made.

| ADR | Fit of the 0431 design | Required evidence or boundary |
| --- | --- | --- |
| [0001](../../../adr/0001-priorities-and-api-layers.md) priorities and API layers | The token is a private low-level capability. PPTX keeps semantic APIs and does not expose ZIP IDs, spans, raw headers, or archive types through ordinary CRUD. Unsupported transfer remains a typed refusal/fallback. | Keep `VerifiedPrecompressedEntry` and source authority below the semantic/public boundary; do not add a raw-byte constructor to a public document API. |
| [0002](../../../adr/0002-crate-topology.md) crate topology | ZIP framing remains in `soapberry-zip`, physical OPC authorization/writing remains in `litchi-opc`, and PPTX owns only the semantic/dependency plan. No facade or peer-format dependency is needed. | Workspace and boundary checks must continue to pass; future token changes must preserve downward ownership. |
| [0003](../../../adr/0003-snapshots-edits-and-patches.md) snapshots and edits | The plan remains tied to immutable source lineage and revision. `current.matches(plan)`, candidate validation, source fences, cancellation, and atomic publication prevent a stale or changed source from authorizing output. | Preserve the typed stale-source and cancellation errors, including checks before allocation, during bounded transfer, after validation, and immediately before publication. |
| [0005](../../../adr/0005-io-memory-and-performance.md) I/O, memory, and measured performance | The design uses caller-owned `ReadAt`, finite hierarchical budgets, bounded decode, explicit cancellation, and a sequential sink. It follows the ADR's raw-copy-when-possible direction while treating the new member as a canonical wrapper. | Charge compressed and decoded storage, wrapper capacity, names, output framing, and checked limits. Keep the accepted result scoped to matched API durations and resource journals; do not promote it to a whole-process, RSS, or physical-copy claim. |
| [0006](../../../adr/0006-validation-security-and-compatibility.md) validation and compatibility | Validation remains non-mutating and precedes output. CRC, decoded size, Deflate termination, descriptor/layout, content type, topology, signatures, active content, and preservation gates remain explicit. Canonical destination wrappers avoid silently copying source framing. | Keep malformed, unsupported, signed, or policy-blocked inputs typed and fail closed. A decoded equality check alone cannot authorize compressed-token reuse. |
| [0010](../../../adr/0010-facade-archive-ownership.md) archive ownership below the facade | No facade regains a `soapberry-zip` dependency or raw archive traversal. The opaque token is consumed by format/package owners. | Do not return the token, ZIP metadata, or archive implementation type through facade signatures or compatibility re-exports. |
| [0011](../../../adr/0011-ooxml-physical-package-ownership.md) OOXML physical package ownership | `litchi-opc` remains the only OOXML physical package boundary. PPTX supplies an authorized semantic payload request; OPC emits the local/central records and typed package errors. | Keep source freshness, signature/active-content policy, destination layout, partial-output accounting, and selected ZIP implementation details in OPC. |
| [0024](../../../adr/0024-current-topology.md) current workspace topology | The applied path follows the current `litchi-pptx` → `litchi-opc` → ZIP ownership layers and does not create an obsolete umbrella owner. | Retain manifest/source ownership checks and do not move physical package logic into a vertical format crate. |

The design review also records two scope limits that are material to the ADR
assessment. The token is not cross-archive raw-member preservation: source
headers, descriptors, timestamps, extras, flags, and offsets are not copied.
And the final precompressed reservation models the captured compressed vector,
generated wrapper capacity, and target-name copies; it is not a whole-process
heap or RSS bound. The accepted API result and portable verification preserve those limits.
