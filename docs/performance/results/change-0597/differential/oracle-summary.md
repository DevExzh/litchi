# Full-corpus `cell()` / `cells()` differential

Three legs of the same transcript over every `.xlsx` under `test-data/`
(180 packages, up to three sheets each, twelve addresses and two ranges per
sheet, a **fresh** `SourceBackedWorkbook` over a counting positional source
for every single read, so each row exercises the cold streaming route).

| leg | what it is | rows | sha256 of the transcript |
| --- | --- | ---: | --- |
| A-base | base 08d968f8e | 4606 | `91e24dc570c235a00dd9c93257ca49bee4dad85af35cfa41be2baadfb1df0831` |
| B-fix | defect fix only (landed) | 4606 | `b56e2785baf153cb58b1b46cd381f0d22f4ac9ec78e45ed060d41de00d501602` |
| C-fix+gate | defect fix + the ineligibility gate (design candidate) | 4606 | `b56e2785baf153cb58b1b46cd381f0d22f4ac9ec78e45ed060d41de00d501602` |

## Leg-to-leg differences

| comparison | rows differing | reads/bytes differing on rows whose result is unchanged |
| --- | ---: | ---: |
| A-base -> B-fix | 434 | 0 |
| B-fix -> C-fix+gate | 0 | 0 |
| A-base -> C-fix+gate | 434 | 0 |

**B-fix and C-fix+gate are byte-identical transcripts**: the ineligibility
gate changes no value, no error, no logical read count and no logical byte
count anywhere in this corpus.

## What the defect fix moved

434 of 4606 rows, over 24 of 180 fixtures. Every one of them
reported the same base value:

- `Err(Invalid("worksheet mergeCells appears before sheetData"))` (434 rows)

What they became (top transitions):

-  204 -> `Ok(Missing)`
-   28 -> `Err(Invalid("worksheet formula expression is empty"))`
-   28 -> `Err(Invalid("invalid worksheet dimension '1:5': invalid A1 range '1:5': invalid or out-of-grid A1 cell reference '1'"))`
-   21 -> `Ok(Stored(Empty))`
-   14 -> `Err(Invalid("shared formula master at (14, 4) is not first in 'D11:D14'"))`
-    8 -> `Ok([SourceCell { address: Cell { row: Row(0), column: Column(1) }, cell: Empty }, SourceCell { address: Cell { row: Row(1), column: Column(1) }, cell:`
-    4 -> `Ok([SourceCell { address: Cell { row: Row(1), column: Column(1) }, cell: Value(Text("SELECT Something")) }, SourceCell { address: Cell { row: Row(1), `
-    4 -> `Ok([SourceCell { address: Cell { row: Row(1), column: Column(1) }, cell: Value(Text("SELECT Something")) }, SourceCell { address: Cell { row: Row(2), `
-    4 -> `Ok(Stored(Value(Text("SELECT Something"))))`
-    3 -> `Ok(Stored(Value(Number(Number("1")))))`
-    3 -> `Ok(Stored(Value(Number(Number("2")))))`
-    2 -> `Ok(Stored(Value(Text(" "))))`

On the 4172 rows whose result is unchanged, reads and bytes are
identical in all three legs. On the changed rows reads rise by exactly nine
per read: the materialized worksheet store the spurious refusal used to
prevent.
