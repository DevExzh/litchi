# Log sections for change 0746

Ready-to-paste paragraphs for the coordinator, one per shared log. This change
does not edit `HOTSPOTS.md`, `REPORT.md` or `GOAL_AUDIT.md` itself.

---

## For `HOTSPOTS.md`

## 0746 — XLS edit owners validate through a validation-only mode of the complete reader; their generic commits stop rendering twice

[0746](0746-xls-validation-only-parse.md) removes the last big item change
[0633](0633-xls-commit-single-framing.md) named for the XLS edit path: every
`cell_values`, comments and sheet-visibility owner opened the complete eager
`Workbook::new` only to validate and read a few facts, building and dropping a
per-cell `BTreeMap` with cloned strings (85% of `Snapshot::from_bytes` on
`54016.xls`). The worksheet walk is now generic over a `CellStore`; the
validation-only store answers the two questions the checks ask of earlier cells
(occupied? latest a `Formula`?) from a banded bitmap and decodes only the cells a
readback keeps, so every record is still validated by the same code in the same
order. The generic commits also publish the rendering `put_stream_shared` had
already validated instead of rendering again (0730's handoff), which was 49% of
the `cell_values` generic commit. Instructions fall 44–51% on every open and
commit touched (37% on `commit_source_backed`), allocations 26–68%, peak live
bytes 44–83%; the public reader is unchanged to the instruction. Wall clock
depends on build layout (a store-forwarding stall on the `CellRecord` copy now
dominates the lean walk): shipped build `54016.xls` open 10.64 → 7.56 ms, plan
11.51 → 4.78 ms, generic commit 24.09 → 13.28 ms; `xls_semantic_one_edit_save`
3.53 → 1.64 ms, `xls_visibility_eager_edit_save` 25.5 → 13.3 ms (before leg built
with the identical command). Remaining:
five SHA-256 overlay fingerprint passes are ~42% of the generic commit; the
measure-only record path would remove the stall.

---

## For `REPORT.md`

## 0746 — XLS validation-only reader mode and validated-render handoff

[0746](0746-xls-validation-only-parse.md) retains five `litchi-xls` commits:
deterministic multi-defect refusals (two hash-ordered refusal sites now name the
lowest offender), a validation-only mode of the complete XLS reader used by the
`cell_values`, comments and visibility owners, and the validated-render handoff
in their three generic commits. Paired ABBA on CPU 20, shipped build:
`54016.xls` `Snapshot::from_bytes` 10.640 → 7.564 ms (0.715),
`commit_source_backed_plan` 11.509 → 4.776 ms (0.415), `commit_source_backed`
14.953 → 11.546 ms (0.768), generic `commit` 24.094 → 13.276 ms (0.553);
`xls-large` open 1.568 → 1.360 ms (0.853) and generic `commit` 3.473 → 2.212 ms
(0.637), with a second build of identical `cell_values` code measuring 0.505 and
0.486 — same instructions, up to 1.7× the cycles, a code-layout effect the record
isolates. Registered selectors against an identically built base harness:
`xls_semantic_one_edit_save` 0.464, `xls_numeric_eager_rk_mulrk_edit_save` 0.516,
`xls_visibility_eager_edit_save` 0.523, `xls_comments_eager_edit_save` 0.792;
source-backed and plan-only
selectors on archive-dominated corpora are flat. Instruction and allocation
reductions are build-independent. Outputs are byte-identical to the base on the
29-case first-error matrix, a 126-fixture census (1,319 rows) and 18,600 mutated
packages. `performance_claim: none`.

---

## For `GOAL_AUDIT.md`

## 0746 — validation work stays, the discarded per-cell model goes

[0746](0746-xls-validation-only-parse.md) keeps every XLS validation obligation
(the complete reopen under default limits, the same parser, coverage, protection,
macro and typed readback checks, the 29-case first-error matrix unchanged) while
removing the per-cell model those owners never read and the second CFB render of
three generic commits. It adds one ADR 0006 determinism fix (orphan-`PtgExp` and
unmatched-comment refusals were hash-ordered). No ADR is amended, no limit
moved, no public API changed, no output byte changed (three-leg census and
mutation differential); the one new failure mode is a typed allocation error
where the old infallible map would abort. Evidence is paired timing,
isolation-pair instructions and deterministic allocation counts; the record
reports that the lean walk's cycles vary up to 1.7× between builds of the same code rather than
a single speedup. No claim registered.
