# 0549 ADR compliance

The accepted ADR/index inventory remains byte-identical to the previously read
inventory, as recorded by adr-manifest.json. The candidate changes one private
CFB collector operation; the common guard is already tracked and unchanged.

| Constraint | Evidence / boundary |
| --- | --- |
| 0001 priorities and layers | Private safe Rust only; exact refusal and bounded resource use precede speed. |
| 0002, 0010, 0011, 0024 ownership | Chain validation stays in CFB. No API or dependency changes in format/facade owners. |
| 0003 immutable snapshots and atomic publication | Only private scratch changes. No document bytes, snapshot, edit, commit or conflict semantics change. Collection still precedes sector claims. |
| 0005 measured performance and budgets | Fresh full matched native/allocator/profile/public-guard campaign. Unchanged reservations, allocation labels and zero-fill. Emitted work must actually decrease. |
| 0006 preservation and safety | Original ordered visited-bit cycle detection remains. The helper checks logical and backing bounds; writing an already-set word preserves its value. Exact errors and scratch reuse/reset are tested. |
| 0008 verification | Reviewed source hashes, exhaustive differential and direct boundary tests, full final quality, per-command custody and strict report replay. No unavailable sanitizer or native-Office campaign is claimed. |
| 0026 OLE directory binding | Physical chain checks, FAT/MiniFAT scratch separation and OLE metadata ownership remain unchanged. |

No ADR exception or reinterpretation is required. Final source retention is
conditional on every frozen performance and quality gate. OLE2/OOXML remain
active; ODF is deferred and iWork excluded.
