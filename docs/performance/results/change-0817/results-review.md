# 0817 terminal admission review

This review covers the terminal independent admission attempt. It does not
relax the frozen plan or authorize qualification, native timing, observer
diagnostics, or workload execution.

## Terminal boundary

`admission-0/receipt.json` records the independent ZIP/XML audit exit code 1.
The retained audit report has SHA-256
`01da3f5b2768ad5ad2a74b4b3b3204f6f72bfadb02e05830e51705dc2d84aa82` and
contains 22 errors. The exporter itself produced five identical policy outputs
for each real case with its own reopen/source gates true; the independent
oracle is the blocking gate required by the protocol.

No qualification, native, or observer result is admissible while this report
is false.

## Real DOCX preservation

The caller-named source is 23,503 bytes with SHA-256
`1cff7a0a94dfce307a70032d21070d26ae34b9fdf742cf70fa66d4a2078ec9d5`. The
default and all four other policy outputs are identical at 23,279 bytes with
SHA-256
`9b36ea76d4bed6dca0d1b567923673d071ca3c2def46a8d0581e7f2d58031e1c`.
The admitted edit is the expected appended marker paragraph in
`word/document.xml`; the audit reported no non-target paragraph semantic
failure.

The only additional changed member is `word/_rels/document.xml.rels`:

- source SHA-256 `f4e76362d55b0a76ff8303c06dc7e1fa4c3447df4992c53af39b77725b47ccb0`,
  817 bytes;
- output SHA-256 `ec271032ef1e480890844bdf7cf60dc5e679ad2f68b3cc076723b4f7dc43336e`,
  815 bytes.

The five relationship tuples are unchanged as a set. The output only moves
the `rId4` `fontTable.xml` relationship from the first position to between
`rId3` and `rId5`; IDs, types, targets, modes, and relationship validity remain
unchanged. This is not relationship loss or a broken OPC graph.

It is nevertheless a frozen-contract failure: the independent oracle compares
the ordered relationship edge list and reports an OPC graph change, while its
decoded-member and canonical-XML comparisons both report the untouched `.rels`
member changed. The preservation rules in `docs/GOAL.md:88-92` require
untouched relationships and ordering to survive, and `docs/GOAL.md:443-447`
requires preservation of ZIP member order and records where the contract
permits it. Whether relationship sibling order is semantically irrelevant does
not make this retained `ok: false` report pass. An order-independent edge-set
comparison alone could not admit this DOCX: the decoded bytes and canonical
XML of the outside-closure `.rels` member still differ. Timing cannot proceed
under the frozen admission rule without an explicit determination and repair of
relationship-order preservation, followed by a fresh complete admission run.

## Real XLSX preservation

The caller-named source is 8,435 bytes with SHA-256
`d7ab3dbb59388d245ee779bf8547748dc6bac70f3c7216e673e0d97dbbbd6bc4`. The
default and all four other policy outputs are identical at 8,521 bytes with
SHA-256
`0f6902152f1b8c40c47086206023887e87f54ef57f1da417bc91d8494a4d3e68`.

The worksheet change is limited to `Munka1!A1`: the source shared-string cell
becomes the ordinary-save marker inline string. Other cell semantics,
worksheet structure, auto-filter, styles, extension data, and relationships
remain unchanged according to the audit. The workbook also has the established
calculation invalidation closure:

- source `calcPr`: `calcId="152511"`;
- output `calcPr`: `calcId="0" fullCalcOnLoad="true" calcCompleted="false"`
  `calcOnSave="true" forceFullCalc="true"`.

Removing the direct `calcPr` element makes the remaining workbook XML equal;
the Rust producer oracle explicitly requires this exact closure for a cell
edit. The producer-side check is
`tools/perf-baseline/src/producer_shape.rs:1420-1460`, and its package
preservation closure names `xl/workbook.xml` and
`xl/worksheets/sheet1.xml` as the expected changed members at
`tools/perf-baseline/src/producer_shape.rs:1463-1496`. The source-backed
regression oracle also asserts the invalidation flags and edited cell in
`crates/litchi-xlsx/tests/source_backed_cell_values.rs:2506-2548`; the generic
cell-edit closure is repeated at `tools/perf-baseline/src/lib.rs:55732-55780`.
The independent audit instead allows only
`xl/sharedStrings.xml` and `xl/worksheets/sheet1.xml`, so it flags
`xl/workbook.xml` as changed outside the target. This is an oracle/closure
mismatch, not evidence of workbook corruption. It still blocks the frozen
admission receipt until the independent contract is reconciled and rerun.

## Outcome

The correct terminal result is **admission blocked; no timing claim**. The
real DOCX graph remains valid and the real XLSX calculation metadata matches
the source-backed edit contract, but neither fact can override the retained
independent-audit failure. A future correction must explicitly determine or
repair DOCX relationship-order preservation; merely sorting relationship edges
in the oracle is insufficient while the untouched decoded member still
differs. It may also reconcile the independent XLSX oracle with the explicitly
verified `calcPr` closure. In either case, preserve this failed receipt and
rerun the complete six-case audit before any timing lane opens.
