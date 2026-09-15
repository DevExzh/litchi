# Log sections for change 0602

Four paragraphs for the coordinator to merge, one each into `HOTSPOTS.md`,
`GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`, in the style of their
newest sections. This record touched none of those files.

---

## For `HOTSPOTS.md`

## 0602 — XLSX real-producer admission is the blocker, not the readback gate

Change 0587's item XLSX-2 ranked widening `stored_entry_is_supported` and the
worksheet-relationship refusal so change 0525's readback reaches Excel-produced
files. Measured on the same 95-fixture corpus, all 95 are refused seven gates
earlier by the package-root relationship allow-list (`snapshot.rs:1924-1952`),
which admits no `docProps`; and `stored_entry_is_supported` refuses none of
them, because all four of its clauses are unreachable — a `t="s"` cell with a
`<v>` fails the planning parse (every parse in `cell_values` passes
`|| Ok(None)`), `<r>` is outside the worksheet element allow-list, and `cm` and
`vm` are outside the `<c>` attribute allow-list, each shown by a synthetic twin.
Independent would-refuse counts: G1 95, G2 77 (no `sharedStrings` in the
workbook allow-list), G3 45, G4 64, G5 0. The mechanism 0525 removes is
nonetheless large, measured at 32.9–85.3% of the commit that pays it on four
derived real-producer fixtures, so XLSX-2's under-10% falsification condition is
not met; but its population is zero. Separately, `verify_authored` refuses the
original bytes of 94 of the 95 fixtures' `workbook.xml` and first worksheet, so
publication blocks real producers even after every editor gate opens. XLSX-2 is
re-scoped from a readback widening to an admission-surface widening with a
`litchi-opc` prerequisite; no production change and no speedup is claimed.
[Design, gates and falsification](0602-xlsx-real-producer-admission-design.md);
[evidence](results/change-0602/README.md).

---

## For `GOAL_AUDIT.md`

## 0602 — real-producer CRUD coverage is blocked before the editor, and before publication

The audit's standing P1 row "cover the high-impact CRUD categories and real
producers" now has a measured obstruction on the XLSX source-backed value path.
The source-backed value editor admits 0 of the 95 real `.xlsx` fixtures: 93
refuse at the package-root relationship allow-list with a typed
`value-only edits refuse package relationship` error naming `docProps/app.xml`
or `docProps/core.xml`, and 2 fail earlier in the read door. Opening that gate
exposes seven more, of which the vocabulary gates are the expensive ones: the
value-only worksheet element allow-list is sixteen names and admits no
`sheetPr`, `mergeCells`, `pageMargins`, `pageSetup`, `autoFilter`,
`conditionalFormatting`, `dataValidations`, `hyperlinks` or `extLst`. Beyond the
editor, the OPC publication audit applies `verify_authored` to the **original**
bytes of every replaced part, and 94 of 95 real fixtures are non-compact
(`FormattingWhitespace`); the only compact package in the corpus is the one
litchi wrote. This confirms as a blocker what change 0587's XLSX section listed
as an item to verify. No harness corpus can exercise any of it: none carries a
shared-string part, a worksheet relationship, an out-of-`sheetData` child or
non-compact XML, and the record specifies the generator shape that would.
`performance_claim: none`; `claim_authorized: false`. OLE2/OOXML remain active;
ODF is deferred until completion and iWork excluded.

---

## For `REPORT.md`

## 0602 — XLSX value-editor admission, sized on derived real-producer fixtures

Four real fixtures were derived onto the value editor's admission surface by two
retained scripts that open the relationship gates and project `workbook.xml` and
each worksheet onto the element and attribute allow-lists while leaving
`<sheetData>` byte-untouched, so the producer's row and cell geometry survives.
Callgrind isolation pairs at N=1 and N=4 with `--separate-callers=1` put one
plan-and-commit operation at 20.9M, 95.4M, 463.9M and 4,110M instructions for 11,
217, 4,108 and 75,770 cells; planning is 47.9–87.5% of that and the complete
worksheet parse 32.3–74.8%. A row-creating `insert` disables change 0525's
reduced readback in-tree, so the same pairs measure the complete candidate parse
directly at 41.7%, 70.5%, 85.3% and 32.9% of the commit that pays it. Paired
A1/B1/B2/A2 latency legs agree: the readback saves 15.4%, 28.2%, 42.3% and 22.1%
of a plan and commit at p50, against an A/A floor measured in the same window of
under 2% on three fixtures and 4.03% on the largest, whose latency numbers are
reported but not relied on. No production code changed and no real producer file
was measured as shipped; the cell geometry is the producer's, the package is not.
See [Change 0602](0602-xlsx-real-producer-admission-design.md);
`performance_claim: none`.

---

## For `ADR_COMPLIANCE.md`

## 0602 — what admitting real producers would cost in contracts

The frozen design widens the value editor's admission surface in five ordered
parts and states the ADR reading for each. ADR 0006's preservation default
argues *for* admission, not against it: a value-only rewrite writes only the
`<dimension ref>` attribute and the cell records between `<sheetData>` and
`</sheetData>`, copying the head and tail byte-verbatim, and in `CT_Worksheet`'s
fixed sequence every `r:id`-bearing child follows `sheetData` while neither
`CT_Row` nor `CT_Cell` has an `r:id` — so a relationship-bearing worksheet is
already preserved by construction and refusing it refuses to preserve what is
preserved. The same argument admits `docProps` at the package root, because the
topology plan names only the worksheet, the workbook and the `calcChain`.
ADR 0005 is not relaxed: every mandated traversal keeps running, complete, over
the complete output, and only the vocabulary the value-only validator admits
widens, to elements the rewrite already copies and never interprets. Admitting
`cm`, `vm` and `<r>` as preserved-but-uninterpreted is what makes
`stored_entry_is_supported` reachable, turning today's dead guard into the
load-bearing safety net it was written to be. Change 0541's six error-precedence
properties survive unchanged and must be re-run with one control per widened
gate, plus two new shared-string orderings. Two doors stay shut: external
relationships, whose closure cannot be proved, and `Error::Signed` on a
signature-bearing package. The publication original-bytes compactness contract
is a `litchi-opc` prerequisite needing change 0528's review and is not designed
here. No ADR is amended and no ADR clarification is proposed by this record.
[Change 0602](0602-xlsx-real-producer-admission-design.md);
`performance_claim: none`.
