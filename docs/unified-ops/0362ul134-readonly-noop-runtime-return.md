# 0362ul134 — readonly no-op runtime return

The fresh runtime packet on frozen candidate
`780cf4c7eebd62784670e2ab14cc49088a4bc91c` passes the first sixteen
common/OPC/DOCX groups. Full-feature PPTX passes 1,398 library tests (one
existing ignored transcript printer) and integration chunks 1–30. Chunk31
passes four readonly tests and the revision test, but fails the later no-op
assertion in `replay_rechecks_readonly_output_with_the_recipient_profile`.
The packet stops as failed; later full/minimal PPTX and outer groups are not
credited as passing.

The physical event-limit refusal added by 126 now succeeds. The failing
assertion instead creates an empty patch from `Package::new()` and applies
it to a recipient with different presentation-properties bytes. The generic
part patch records its originating package even when empty. Its complete
content-base comparison correctly returns `PatchConflict` before semantic
validation. The recorded-base contract requires this refusal; production
scope must not be weakened to satisfy the test.

Return134a corrects the consumer to create its no-op from the matching
recipient, retain same-owner byte preservation, exercise an equal-content
recipient with tighter limits, and explicitly reject an unrelated pristine
package without mutation. Independent source review confirms the behavior;
an unused `mut` is returned for removal before composition. Analogous math,
metadata and laser consumers already use matching bases. No production
change or repaired-source runtime pass is claimed.

Root independently verifies the terminal capture, source and all three
approved locks, exact commands, logs/traces, process exit and target bounds.
Evidence is external under the worktree parent's
`(deleted-coord-artifact)`:

| File | Bytes | SHA-256 |
|---|---:|---|
| `rows.json` | 73,691 | `992693992302ce1c3e69bf0c698be460e998e64ec6177924fc85fc537c950ade` |
| `result.json` | 78,244 | `8f92b4993e0303ab8f812174ecb1c9a3d0f9ed40fad3ed9f794196dd45e7e273` |
| `pptx-full-tests/0031.log` | 216,837 | `224d0684adbea9c9a4d4360ef571e359454e0a49014d1b6e017b88d2638425b9` |

The scoped reproduction is the packet's captured chunk31 command, under
the approved F5 supervisor and the same source/locks; the complete packet
uses `python3 -B run_f5.py --go` with fresh state/target directories. Root
terminal verification uses `(deleted-coord-artifact) STATE`.
The failed capture remains unchanged. Remaining coverage is prepared as a
separate diagnostic continuation. Glossary26 starts on the same frozen C
after root preflight; its result remains pending.

## Remaining qualification engineering

The raw-census v9 adapter passes root and independent source review after
fixing initial binding/manifest sealing and a reproduced recursion defect
in v8. Nine build-free tests and injected capture/finalization paths pass.
The C-bound packet remains unexecuted. Old generated source overlays must
be inventoried and cleared after active validation, preserving non-generated
evidence and the three locks. The execution plan must describe the actual
packet root, and Python children must inherit bytecode suppression.

The OfficeCLI postcheck is prepared for the existing full53 CLI row, with
the exact 96-script and 6,808-row baseline. It does not require a duplicate
supervisor run. Its source binding will be refreshed if composition changes
the candidate.

A source audit confirms that native policies 72/74/75 are not yet executable
v2 qualification paths. DrawingML's paired consumer still uses v1 receipts;
signed-native preparation/execution remains disabled; current-native lacks
the complete three-role lock and legacy execution path. These are required
engineering work. Separate protocol leaves 137, 136 and 135 respectively
implement the consumers and meaningful tests. Their trusted final source
pins belong to a separate tooling checkout, avoiding a source commit that
tries to pin its own hash. No pending manifest or historical receipt is
accepted as a v2 measurement.

All new leaves remain separate until reviewed composition. Final workspace,
API census, fidelity/performance and trunk integration remain open. Root
stays on `feat/unified-ops-api`; nothing is pushed.
