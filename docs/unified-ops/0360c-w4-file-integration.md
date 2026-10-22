# 0360c — Assemble W4 with the File options candidate

- Inputs: W8 Part A `fd5f513df6` and W4 `b6a5d10935`.
- Branch: `uo2/coordinator-w4-integration`.
- Status: integration candidate; performance acceptance and trunk merge pending.

The coordinator combines the pinned W4 API migration with Part A's File
options, required detector policies, ODF semantic ingress and validated File
measurement harness. The integration keeps the PPTX shared immutable OPC
owner while adopting the OPC result published by the shared metadata engine.
It preserves both sides' benchmark-matrix entries. The source-content-types
presence accessor from W4 is retained at this assembly point; the later XLSB
integration removes its superseded form.

The all-target check found an unmigrated facade sparkline example. Its enum
imports and uses now follow the shared `AxisType`, `EmptyCells` and
`SparklineType` names. The feature matrix also reproduced the iWork-only
detector's unused policy parameter. The already validated flat-ODF fix marks
that parameter consumed only in feature combinations whose ZIP-tail probe
owns the format-specific policy. The public required-policy API is retained.

The [validation receipt](results/0360c/validation.json) verifies every
external command log, including failed attempts. Nine affected packages pass
all-target/all-feature compilation, and 119 PPTX metadata tests pass. The
feature matrix has 27 initial successes and all four previously failing
iWork cases pass after the cfg-only repair. Boundary, benchmark-matrix,
workspace formatting, locked all-target harness compilation and harness
formatting checks pass. The 77 incoming wave commits contain no pre-rewrite
commit IDs. The initial copied lockfile was insufficient for the
combined dependency graph; offline metadata resolution refreshed the ignored
lockfile before compilation. Its initial locked failures and the initial
example/feature failures are retained in the external logs. The harness
also needed its tracked lockfile reconciled with the already required
`siphasher` 1.0.4 dependency, and six detector call sites migrated to the
required policy argument. The timed detector routes reuse an explicit default
policy prepared outside their measured calls. Three harness files receive
formatting-only cleanup; their earlier formatting failure is retained.

W4's later source-owner policy compaction, unresolved owner-admission memory
cost and no-op metadata base checks are not accepted by this pinned assembly.
XLSB `b2e21c0ab4`, the remaining W8 branches and W5 repairs still need
integration and final source qualification. Review additionally found DOCX
and PPTX `apply_metadata` returning a Package instead of the required Commit;
their consumer migration and the gate/inventory coverage gap remain required.
The DOCX source-backed primary replay and PPTX primary/source-backed replay
also return `Self`; the final sweep must audit every owner replay result and
shared triad type against the contract.
The final full workspace gate and owner approval to push remain outstanding.
