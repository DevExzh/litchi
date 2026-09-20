# Reference metadata native receipt

`reference-metadata-native.fods` is the retained local LibreOffice fixture for
the eight reference and worksheet metadata functions in this batch.  It uses
four ordered sheets (`Main`, `Data`, `Archive`, and hidden `Hidden`); `Hidden`
has a table-family automatic style with `table:display="false"` so the
hidden-sheet claim is represented in the input itself.  The fixture covers
single references, rectangular references, 3-D references, ordered and
duplicate lists, omitted current-position calls, scalar axis publication,
matrix axis output, arrays, and pseudotype refusals.  The formula rows perform
metadata operations only; they do not depend on cell contents.

`recalculated.ods` was produced with `/usr/bin/libreoffice` 26.2.5.2 using a
fresh headless profile and `C.UTF-8`.  `reproduce.py` repeats the conversion in
a new temporary profile, checks the retained input, output `content.xml`, and
typed formula-row hashes, verifies that the `Hidden` table still has
`table:display="false"`, and removes its temporary profile and output tree.
The hashes and converter details are recorded in `provenance.json`.

The retained run has 32 formula rows: 28 match the independent local profile
and four are explicit host divergences.  LibreOffice reports `Err:504` for
`COLUMNS`, `ROWS`, `SHEET`, and `SHEETS` when given a multi-entry
`ReferenceList`; the local contract records those known pseudotype refusals as
formula `#VALUE!`.  These host tokens are retained verbatim and are not
translated into normative results.

An additional fresh-profile probe used while resolving overload behavior found
that LibreOffice reports `Err:504` for `SHEET(1)`, `SHEET(TRUE())`,
`SHEET({1;2})`, and `SHEETS({1;2})`, and returns `FALSE` for `ISREF({1;2})`.
Scalar `COLUMN([.B2:.D4])` and `ROW([.B2:.D4])` both returned `2`.  Those
observations informed review only; the normative Number/Logical conversion,
array handling, and scalar first-axis rules come from `contract.md`, not from
LibreOffice compatibility behavior.
