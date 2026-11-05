# 0362ul94 — shared OOXML publication composition and error return

This private composition starts at integration
`88ca2fd38fdd23ca2ee4ad22449a83893804e4e1`. It joins reviewed Web/Ribbon
publication `ce75a70582d089a13ee35e330d5e1f88b8f6d5e9` and VBA publication
`b64d3b55287c22fc9e48619309f3f83bc3a01c7d` at merge
`6bf2a5f68a`. Source composition is not runtime or performance acceptance.
Properties publication is separately under review and has not been merged
into this candidate.

The common glossary edits merge without a conflict. The Web/Ribbon checker
return uses only the explicit audited binding from
`litchi_opc::SourceLimits` to `litchi_opc::limits::SourceLimits`;
wrong canonical modules and ambiguous identities remain rejected. Web patch
no-op inspection is `is_noop`, with every former `is_empty` caller migrated.
No compatibility alias remains.

## Preserved compiler and glossary evidence

Web/Ribbon actual state
`(deleted-coord-artifact)`
passed rows 0–8: formatting, common/four-format/ops/facade compiler checks,
and their Clippy checks. Row 9 failed the common-minimal glossary profile
with six findings. Five were return/nominal-identity checker errors; the
sixth exposed Web Patch's missing `is_noop`. The source return fixes all
six, but a fresh rustdoc capture and runtime tests are still required.

VBA actual state
`(deleted-coord-artifact)`
passed formatting, then failed the common compiler row with 12 OPC errors.
The reviewed `b64d3b5528` repair preserves namespace resolution without
allocation, deferred fixed-buffer comparisons, source spans and lexical
fidelity. No later compiler or runtime row is claimed as passed.

| State artifact | Bytes | SHA-256 |
|---|---:|---|
| Web `rows.json` | 39,243 | `00d3ff25df5f2bacae720d4a468498dacf65d41231879915fbc8a4387d2257ec` |
| Web `result.json` | 41,942 | `8c1edf088c78367e3629c8e593730c2e73de5abd17db2c4de30a11babe2cf664` |
| VBA `rows.json` | 7,531 | `fc565f7919e6b362e73304e28e4cb45d54856aee587cca415530f45340e1cbcb` |
| VBA `result.json` | 8,463 | `9a87300ca548b7ed4c1a4ad34ad1d6e0b6a7241391d6784319c46acd2d8270cb` |

## Required error-contract return

Independent review found visible classification defects. Web/Ribbon/VBA
used generic package errors for stale recorded bases, signed mutations,
valid unsupported requests and generated-state invariants. Unified
classifiers consequently reported some of these as `corrupt_file`.

The coherent follow-up must add crate-owned typed common errors and migrate
all producers, format adapters and unified classifiers in the same commit:

- recorded-base mismatch: `PatchConflict`, classified `invalid_input`;
- source freshness change: typed `SourceChanged`, classified `io_error`;
- unsupported but well-formed requests: `Refused(Refusal)` with the correct
  kind and a fixed, content-free reason;
- signed changed publication: the signed-write error;
- violated generated-state invariants: internal error;
- malformed retained graph/XML: corruption, preserving existing distinctions.

Properties has the same obligations in its separate local error type,
including metadata consumers. These are required repairs under 0351 §2.4,
not accepted exceptions. Actual isolated validation of the resulting combined
source, final inventory, corpus/native and performance qualification remain
open. No push is authorized by this record.
