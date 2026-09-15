# Log sections for change 0599

Four paragraphs for the coordinator to merge, one per log, in the style of each
log's newest sections. Nothing here edits `HOTSPOTS.md`, `GOAL_AUDIT.md`,
`REPORT.md` or `ADR_COMPLIANCE.md` directly.

---

## For `HOTSPOTS.md`

### 0599 — XLSB cell-value commit: one workbook parse instead of two, none for a proven no-op

Survey item XLSB-1 (rank 19 of change 0587) is implemented in full, and it is
the first XLSB *performance* record in the program — all sixteen prior XLSB
records (0304–0376) are correctness records. `Workbook::apply_cell_values`
cloned its package and reparsed the entire eager `Workbook` unconditionally,
including when the callee had just proved the patch reproduced the stored
worksheet bytes exactly; a real edit took **three** `OpcPackage` clones, one
worksheet-blob copy and **two** independent full parses of byte-identical
candidate bytes, discarding the first after validating against it. The callee is
now a crate-private `apply_retaining_parse` returning `Applied::Unchanged` or
`Applied::Published{snapshot, workbook}`: a proven no-op clones nothing and
parses nothing, and a real edit publishes the very parse that
`validate_dependencies` ran against. Instructions per commit fall
**52,308,801 → 29,325,884 (−43.94%)** for `noop_transaction_commit_save` and
**78,314,813 → 54,162,314 (−30.84%)** for `edit_one_existing_scalar_save` on
`testVarious.xlsb`. The per-symbol difference is *identical to the instruction*
between the two cases, which is the signature of exactly one
`from_opc_package_with_external_link_limits` removed in each; the 1,169,582 Ir by
which they differ is the edit path's two extra clones and its blob copy, and
package clones are cheap only because `BlobPart` holds `Arc<Vec<u8>>`. Paired
timing (A1 B1 B2 A2 plus an A/A pair, 40 samples a leg, CPU 17) gives p50
**−46.43%**, **−31.40%** and **−29.38%** for the three commit cases on
`testVarious.xlsb` and −9.69%, −9.69%, −8.29% on `cond_format.xlsb`. Every commit
case moved down on every fixture and nothing regressed above the review
threshold: the largest positive p50 delta in the 32 case×fixture cells is +1.93%,
on an untouched read path. The saving tracks a workbook's **feature
surface**, not its size: it is the cost of one eager parse, dominated by the XML
of drawings, pivot caches, tables, chart sheets and connections, so it is 44% of
a no-op commit on the feature-rich 22.7 KB real producer and 7% on a
feature-free 324 KB synthetic. Change 0587 named the corpus ceiling — 22,715
bytes, one sheet, 48 cells — as the blocker for every XLSB scaling question;
this change lifts it with a checked-in generator
(`tools/perf-baseline/src/bin/xlsb_synthetic_fixture.rs`) rather than a
checked-in blob, and the 4-sheet, 90,000-cell fixture it produces is what shows
that the saving does **not** scale with cell count. What remains ranked in this
area: `cell_values::root`'s two sibling sites still reparse a second time and
have no harness selector; XLSB-2 stays frozen behind 0581's ADR; XLSB-3 stays
ADR 0005-blocked; and a **large real-producer `.xlsb`** is still absent, which
is now the one measurement that would turn this change's 7%–44% range into a
number.
[Change 0599](0599-xlsb-commit-single-parse.md);
[evidence](results/change-0599/README.md); `performance_claim: none`.

---

## For `GOAL_AUDIT.md`

| priority | item | what it needs |
| --- | --- | --- |
| P1 (progressed) | Finish source-backed CRUD adoption across formats — the XLSB publication half | Change 0599 removes the duplicated whole-workbook readback from the XLSB cell-value commit and skips it entirely for a proven no-op, the first XLSB performance evidence in the program. What stays open is unchanged by it: XLSB cell reads still go through the eager `OpcPackage` door (0581's frozen ADR question), `SourceBackedWorkbook` still feeds only a sequential text writer, and the publication path for `apply_workbook_structure`, `apply_sparklines` and `apply_cell_watches` is unmeasured because **no harness selector opens it**. Registering one is the prerequisite for finishing the pattern in this crate. |
| P1 (progressed) | Cover the high-impact CRUD categories and real producers — XLSB corpus ceiling | The ceiling change 0587 recorded (largest `.xlsb` 22,715 bytes, one sheet, 48 stored cells, so a one-cell read and a full scan are 0.3% apart) is now liftable on demand: `tools/perf-baseline`'s `xlsb_synthetic_fixture` generates a deterministic workbook of any shape through the public writer and proves it reopens before writing it. Change 0599 measured 4-sheet fixtures of 15,000 and 90,000 cells with it. The gap it does **not** close is a *real-producer* large `.xlsb`: the generator emits no pivot cache, table, chart sheet, drawing, connection, external link or VBA project, which is exactly the surface whose parsing dominates the figures. Change 0599's saving therefore has a measured range, 7%–44%, and no single number. |

Supporting note for the audit body: change 0599 is a GOAL step 1 result —
*eliminate unnecessary work* — and its evidence separates the two methods rather
than leaning on either. The five `xlsb_crud` read-only cases never reach
`apply_cell_values`, so their paired delta (−3.54% to +1.93%) bounds drift and
code layout in the same window; subtracting that offset from the no-op delta
gives about −43% against the deterministic −43.94% of instructions. That window's
A/A floor was **worse than the host's standing figure** — |p50| median 0.84%,
p90 3.22%, max 10.04% over 64 control pairs — and the record respects it: the two
synthetic fixtures' commit deltas fall inside it and **nothing is claimed from
their timing**. No allocation, RSS, syscall, cold-cache, physical-device or
cross-platform measurement was taken; three allocations per commit are removed by
construction and none was counted, because `xlsb_crud` still has no
allocator-metrics hookup.

---

## For `REPORT.md`

Change 0599 makes an XLSB cell-value publication parse its candidate workbook
once. `Workbook::apply_cell_values` previously cloned its package before knowing
whether anything would change and re-derived the complete eager `Workbook`
afterwards in every case, and the inner
`cell_values::workbook::apply_with_external_link_limits` separately cloned twice
more and built a throwaway parse purely to run `validate_dependencies` against
it. The publication boundary is now a crate-private `apply_retaining_parse` that
takes `&OpcPackage`, answers an exact no-op from the stored bytes without
cloning or parsing anything, and otherwise hands back the validated candidate as
`Applied::Published` for the caller to install. `apply` and
`apply_with_external_link_limits` keep their signatures, their `&mut OpcPackage`
publication contract and their error identity, so the two in-crate callers in
`cell_values/root.rs` are untouched. Validation is unchanged and still runs
exactly once: `require_worksheet`, the patch application against the stored
bytes, the complete candidate reparse that is the whole-workbook readback, the
candidate worksheet-URI lookup and its `WorksheetNotFound` refusal, the typed
worksheet decode, and `validate_dependencies` over every committed cell's style,
font, fill, border, number format, shared-string index and rich-string run fonts.
A 17-artifact, 35-worksheet differential between the two legs — worksheet
catalogs, per-worksheet snapshot and cell digests, an exact no-op and a real
scalar edit per worksheet with published and saved digests and a readback, and a
refused publication per artifact — is **byte-identical**, and `xlsb_crud`'s own
gates are `true` in all 192 observations, with the changed-member list empty for
every no-op and exactly one worksheet part for every edit. Measured on
`testVarious.xlsb`: instructions per commit −43.94% (no-op) and −30.84% (one
cell), paired p50 −46.43% and −31.40%; on `cond_format.xlsb` −9.69% and −9.69%.
The two synthetic fixtures' deltas lie inside this window's A/A floor and carry
no timing statement. `performance_claim: none`; no speedup, regression, RSS,
allocation, cold-cache or cross-platform claim follows. See
[Change 0599](0599-xlsb-commit-single-parse.md);
[evidence](results/change-0599/README.md).

---

## For `ADR_COMPLIANCE.md`

Change 0599 tightens, rather than relaxes, the ADR 0003 publication boundary for
XLSB cell-value commits. ADR 0003 requires that "public format editors publish
only after their staged CRUD operation and typed readback succeed"; the new
`apply_retaining_parse` takes the package by shared reference and builds its
candidate in a local, so every refusal on the way — the missing or non-worksheet
part, the stale patch, the candidate parse, the candidate worksheet-URI lookup,
the typed worksheet decode, and `validate_dependencies` — returns before anything
the caller owns has been touched, and the single assignment that publishes is the
last statement executed. The old shape mutated the callee's `&mut OpcPackage`
argument one statement earlier. Two properties are asserted rather than argued:
an in-crate test destructures `Workbook` **exhaustively** — so a field added later
will not compile until the projection covers it — and proves that after a real
edit the published workbook equals a fresh
`from_opc_package_with_external_link_limits` over the bytes it published, and
that after an exact no-op the workbook is both unchanged and still equal to a
fresh parse of its own package; a third test proves a refused publication leaves
that projection identical. At the public boundary, two further tests prove that
an exact no-op publication and a refused publication each save a **byte-identical
package**, which is the ADR 0006 preservation obligation, and the 17-artifact
differential extends that to every `.xlsb` in the corpus, including the exact
typed refusal string. No public signature changed; the two new items are
`pub(crate)`. No `unsafe` was added, no limit relaxed and no malformed-input
defence weakened: `entries.try_reserve(1)` per cell, `Cursor::guard`,
`raw::record::Limits`, `cell_values::Limits`, `litchi_opc::ReadLimits` and
`ExternalLinkLimits` are untouched, and the harness re-proves the last three
refuse on every fixture in every leg. One question is raised and **not** answered:
`Workbook::apply_sparklines` and `Workbook::apply_cell_watches` publish straight
into `&mut self.package` and never refresh the derived workbook fields at all,
which is a different shape from the boundary this change tightened. See
[Change 0599](0599-xlsb-commit-single-parse.md); `performance_claim: none`.
