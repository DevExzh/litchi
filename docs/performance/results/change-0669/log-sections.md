# Log sections for change 0669

## For `HOTSPOTS.md`

## 0669 — XLSB edit owners publish the parse that validated the candidate

Record: [0669](../../0669-xlsb-edit-residues.md). The remaining row-13 XLSB
residue is implemented. `insert_candidate_cell` and `transfer_cell` now install
the `Workbook` returned by the existing single-parse cell-value publication
seam, removing their second candidate reparse. Sparkline and cell-watch
publication use the same retained-parse outcome, so their facade owners replace
all derived fields with the parse that validated the changed worksheet. The two
resource paths independently ensure the part and its exact workbook
relationship, repairing an existing part whose relationship is missing. The
two resource witnesses and the successful/stale facade tests pass. No timing
claim is made; 0674 owns the `apply_workbook_structure` selector and gates.

## For `GOAL_AUDIT.md`

## 0669 — the XLSB publication boundary remains staged and bounded

Record: [0669](../../0669-xlsb-edit-residues.md). The change follows the goal's
correctness-first rule. Candidate bytes are patched on a cloned package,
decoded, unsigned when changed, and parsed as a complete workbook before one
final owner assignment. Any patch, worksheet, dependency, or workbook error
leaves the caller unchanged; exact no-ops return without a candidate parse.
Installing the retained parse refreshes the workbook-derived fields without
relaxing a limit or bypassing readback. Resource relationship repair occurs on
the staged transfer package and matches internal mode, exact type, and exact
target, so a dangling or external relationship cannot authorize a missing
resource. The test evidence covers both repaired resource states and the
successful and stale facade paths. No speed figure is inferred from this
correctness evidence.

## For `REPORT.md`

## 0669 — the named XLSB edit residues are closed

Record: [0669](../../0669-xlsb-edit-residues.md). A workbook-structure replay
that inserts or transfers a cell no longer reparses the candidate after the
candidate parse has already validated it. Sparkline and cell-watch edits now
publish the complete workbook built from their candidate, keeping package bytes
and derived metadata in one transaction. Styles and shared strings also repair
the case where the part is present but the workbook relationship is absent.
`cargo fmt`, `cargo check`, clippy with warnings denied, all 575 library tests,
and the 144 XLSB integration tests pass with the requested jobs-2 and debug
settings. This record reports no benchmark or registered performance claim;
the selector needed to price the remaining path is handed to 0674.

## For `ADR_COMPLIANCE.md`

## 0669 — retained parse publication preserves ADR 0003, 0005, and 0006

Record: [0669](../../0669-xlsb-edit-residues.md). ADR 0003's source-checked,
atomic publication is preserved by borrowing the published package, staging the
candidate, and assigning the validated workbook only at the end. ADR 0005's
typed validation and finite resource limits remain in place; no check is
removed, and relationship repair is limited to an already staged candidate.
ADR 0006's preservation default remains exact for no-ops and validated for
changes. The accepted lazy OPC ADR 0030 is not amended and no proposed ADR is
cited by production code. No unsafe code, dependency, global cache, ambient I/O,
executor, or archive ownership leaks into the public API.
