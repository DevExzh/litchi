# Log sections for change 0746

Ready-to-paste paragraphs for the coordinator, one per shared log. This change
does not edit `HOTSPOTS.md`, `REPORT.md` or `GOAL_AUDIT.md` itself. Revised
after the independent review; the numbers are the record's.

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
(occupied? latest a `Formula`?) from an occupancy map of 64 bytes per occupied
row and decodes only the cells a readback keeps, so every record is still
validated by the same code in the same order. The generic commits also publish
the rendering `put_stream_shared` had already validated instead of rendering
again (0730's handoff), which was 49% of the `cell_values` generic commit.
Reviewed build: instructions fall 43.9–51.0% on every open and commit touched
(37.0% / 41.0% on `commit_source_backed`), allocations 26–68%, peak live bytes
44–83%; the public reader is unchanged to the instruction. The review found that
the first occupancy map (256-row bands, 16 KiB each) made a crafted sparse
workbook, one cell per band, slower than the base; the final map is proportional
to occupied rows, and in one window the final source runs that shape at
0.433–0.453× the base, the `54016.xls` open at 10.46 → 4.87 ms (0.465) and the
generic commit at 23.24 → 10.12 ms (0.436), at 4.1–9.6% more allocated bytes
than the banded map on `54016.xls`. Wall clock of this lean walk varies with
build layout and host state (one binary's cycles per `54016.xls` open measured
29.99 M and 21.98 M four minutes apart, instructions unchanged); instructions
and allocations are the stable evidence.
Remaining: five SHA-256 overlay fingerprint passes are ~42% of the generic
commit; the measure-only record path would remove a store-forwarding stall.

---

## For `REPORT.md`

## 0746 — XLS validation-only reader mode and validated-render handoff

[0746](0746-xls-validation-only-parse.md) retains five `litchi-xls` commits and
four review-fix commits: deterministic multi-defect refusals (two hash-ordered
refusal sites now name the lowest offender), a validation-only mode of the
complete XLS reader used by the `cell_values`, comments and visibility owners,
and the validated-render handoff in their three generic commits. Paired ABBA on
CPU 20, reviewed build (`b85d3e534c`): `54016.xls` `Snapshot::from_bytes`
10.640 → 7.564 ms (0.715), `commit_source_backed_plan` 11.509 → 4.776 ms
(0.415), `commit_source_backed` 14.953 → 11.546 ms (0.768), generic `commit`
24.094 → 13.276 ms (0.553); a second build of identical code measured up to
1.7× fewer cycles with the same instructions. Registered selectors against an
identically built base harness: `xls_semantic_one_edit_save` 0.464,
`xls_numeric_eager_rk_mulrk_edit_save` 0.516, `xls_visibility_eager_edit_save`
0.523, `xls_comments_eager_edit_save` 0.792; source-backed and plan-only
selectors on archive-dominated corpora are flat. After review, the occupancy
map is sized by occupied rows instead of 256-row bands, which had made a
crafted sparse-band workbook 1.47–1.58× the base's instructions; final source
against the base in one window: sparse-band opens 0.433–0.453, `54016.xls` open
0.465, generic commit 0.436, `xls-large` open 0.444, public reader flat.
Outputs are byte-identical to the base on the 29-case first-error matrix, a
126-fixture census (1,319 rows) and the opens of 18,600 mutated packages, on
both the reviewed and the final source, and a new frozen 27-case
worksheet-level multi-defect matrix generated on the base passes in all three
store configurations. `performance_claim: none`.

---

## For `GOAL_AUDIT.md`

## 0746 — validation work stays, the discarded per-cell model goes

[0746](0746-xls-validation-only-parse.md) keeps every XLS validation obligation
(the complete reopen under default limits, the same parser, coverage,
protection, macro and typed readback checks, the 29-case first-error matrix
unchanged, and a new frozen worksheet-level multi-defect matrix whose
expectations were generated on the base) while removing the per-cell model those
owners never read and the second CFB render of three generic commits. It adds
one ADR 0006 determinism fix (orphan-`PtgExp` and unmatched-comment refusals
were hash-ordered; they now name the lowest position or object id). No ADR is
amended, no limit moved, no public API changed, no output byte changed
(three-leg census and mutation differential, rerun on the final source). The
only new failure mode is resource exhaustion. Inside a worksheet it fails that
worksheet's parse with a typed `Allocation` error, the package walk drops the
sheet as it drops any worksheet refusal, and the edit owners refuse with their
typed coverage `UnsafeEdit` where the old infallible map would abort; listing
the few cells a readback keeps returns the `Allocation` error directly. The
occupancy map is proportional to the occupied rows (the independent review's
sparse-band case is fixed). Evidence is paired timing, isolation-pair
instructions and deterministic allocation counts; the record reports that the
lean walk's cycles vary with build and host rather than a single speedup. No
claim registered.
