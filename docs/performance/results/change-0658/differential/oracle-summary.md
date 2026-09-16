# Differential summary: change 0658

## 1. Corpus transcript (the oracle)

Every `.xlsx` under `test-data/` (180 packages), up to three sheets each, twelve
addresses and two ranges per sheet, through the public `SourceWorksheet::cell`
and `SourceWorksheet::cells`, with a **fresh** `SourceBackedWorkbook` over a
counting positional source for every single read, so every row exercises the
cold streaming route. Each row records the exact `Result` debug, the logical
`read_at` count and the logical bytes read.

| leg | lines | read rows | sheet-catalog lines | sha256 |
| --- | ---: | ---: | ---: | --- |
| before (`70d7768cc`) | 4786 | 4606 | 180 | `23abae8318fb74c5c38a1ddc3d70f662b16b465ca8e7763abec2e89a906b227a` |
| after (this change) | 4786 | 4606 | 180 | `23abae8318fb74c5c38a1ddc3d70f662b16b465ca8e7763abec2e89a906b227a` |

The two transcripts are **byte-identical**: same values, same refusals (type and
text), same fallback outcomes, same logical reads, same logical bytes. The gate
removes XML parsing, not source I/O — the verified OPC reader still drains and
CRC/size-verifies the whole member on both legs.

The transcripts themselves (about 1 MB each) are not retained; the sha256 sums
above identify them and `../probe/` regenerates them deterministically
(`probe0658 oracle test-data`).

## 2. Verdict census (`verdicts-before.tsv`, `verdicts-after.tsv`)

The first three worksheet parts of every package, extracted with `zipfile` and
read through `litchi_xlsx::raw::selected_worksheet::scan` over a counting
`BufRead`. Columns: part, verdict, bytes pulled from the reader, part size.

- **326 parts, all of them ineligible** — 325 `UnsupportedStructure`,
  1 `RichInlineText` — and every verdict is identical on both legs.
- Bytes pulled from the worksheet reader, aggregated: **9,324,342 → 795,708
  (-91.47%)** of 9,324,342 part bytes.
- The counter is quantized to the stream's 8 KiB fill, so the 288 parts smaller
  than one fill still show 100%. Every one of the 38 larger parts pulls exactly
  one 8,192-byte fill after the change and no more, which is 0.24% of the
  largest part and 89.54% of the smallest of them.

## 3. First-error matrix (`matrix-before.tsv`, `matrix-after.tsv`)

`make_matrix.py` (change 0597's generator, retargeted to this repository's
tracked fixture) writes 60 synthetic packages: one small valid package's
`xl/worksheets/sheet1.xml` rewritten with a worksheet malformed — or merely
ineligible — at a chosen position, in a `<cols>`-bearing shape and a
`<cols>`-free control shape. Each package is read at 15 coordinates through the
same public path: 900 rows per leg.

- **210 rows over 15 of the 60 packages move**; 690 rows over 45 packages are
  identical.
- Every moved row is a package that is both ineligible and malformed **after**
  its mark. `matrix-witnesses.txt` is the complete before/after list; the change
  record tabulates the eight distinct witnesses and their classes.
- The generated packages are not retained; `make_matrix.py` regenerates them
  deterministically.
