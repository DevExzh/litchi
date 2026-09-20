# Lookup-function native receipt

`lookup-functions-native.fods` is a bounded LibreOffice probe for all nine
lookup/reference functions: ADDRESS, CHOOSE, HLOOKUP, INDEX, INDIRECT, LOOKUP,
MATCH, OFFSET and VLOOKUP. It retains a four-sheet workbook (`Main`, `Data`,
hidden `Hidden`, and `Archive`), duplicate ascending keys, an Empty search
seam, formula-error cells, a 3-D reference, list refusals, and inline arrays.
The array spellings use OpenFormula's semicolon-within-row and pipe-between-row
grammar. The Hidden table has a table-family style with
`table:display="false"` in the input and recalculated output.

The capture used `/usr/bin/libreoffice` 26.2.5.2 (Build 620), headless with a
fresh profile and `C.UTF-8`. `recalculated.ods` and its extracted
`content.xml` are retained. `native-results.json` preserves each typed native
cell result and records the independent contract expectation alongside a
`parity` or `native-divergence` disposition. The capture has 32 retained
formula rows: 30 function observations (17 contract parities and 13 explicit
LibreOffice divergences) plus two Data-sheet formula cells that supply the
error seam used by the probe.
Host `Err:504`, range-display `#VALUE!`, lack of Unicode sharp-s folding, lack
of Empty-to-zero lookup normalization, and native formula-error precedence are
retained verbatim; they are not translated into normative results.

Run the reproducible receipt with:

```text
python3 reproduce.py
```

The script converts the FODS in a new temporary profile, checks the input and
Hidden-table shape, compares typed formula rows and `content.xml` against the
retained capture, then lets the temporary profile tree disappear. It does not
replace retained files.

Retained hashes:

| File | SHA-256 |
| --- | --- |
| `lookup-functions-native.fods` | `f084909cae337df25c37e028fa4644eac5c80a68b78522cc6c7ebbb2e592cb60` |
| `recalculated.ods` | `4a6eb73fb895fffda6654a07a94a77e894210228ef7f4c10943eebbadfc05ae4` |
| `content.xml` | `ecd94bdfc0acfff99eb3890148adaf7267243d11ae2543b987217706fc4a9175` |
| `native-results.json` | `ecb3f972895a260f95fdfc0285e881d4cea872c701e3ec0752901cfcede4fbb1` |

The independent oracle is bound to contract SHA-256
`b112d66d687337912333f932c199f6e0ee5241fedc335aeaa95de572889e6aaf`; its
script and 127-row goldens hashes are recorded in `provenance.json`.
