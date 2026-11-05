# Record 0015 evidence: leaf write costs (throwaway probe)

These are the leaf measurements behind record 0015 §5.6. They time leaf APIs
only; nothing here measures the unified layer, which does not exist yet.

## Method

- Tree: `uo/0015-m2-plan` at `a8ba410287`, toolchain 1.95.0.
- Probe: a standalone Cargo package in the session scratchpad with path
  dependencies on the tree's leaf crates, locked to
  `tools/perf-baseline/Cargo.lock` (the worktree keeps no root lock file;
  every package kept its locked version), release profile with the
  workspace's `lto = true`, `panic = "abort"`. Built with
  `CARGO_TARGET_DIR=<worktree>/target CARGO_INCREMENTAL=0 cargo build --release`.
  The probe sources were throwaway and deleted; no binary is committed.
- Host: AMD EPYC 9R45, 32 cores, shared with other jobs; every measurement
  pinned with `taskset -c 7`. Load averages during the runs: 2.2–8.5.
- Each data point: one calibration run, then 11–201 timed repetitions after
  1–19 warm-ups (201 for points under 0.2 ms), each from the same base: the
  XLSX eager points reuse one immutable `Workbook`; the DOCX, PPTX and XLSX
  source-backed points open a fresh owner from the same `Arc<Vec<u8>>`, timed
  as its own phase. Output goes to a pre-sized `Vec`.
- Values are medians (p50) with min–max, in µs, rounded to three significant
  figures; `<1` is below 1 µs. Millisecond points: max−min ≤ 15% of p50.
  Points repeated in separate runs agreed within about 1–2%.
- Every X4, D1 and P1 output was re-opened and the edited values read back
  once per data point.
- Single host, single session, no ABBA: these ranked engines for the plan;
  they are not budgets or claims.

## Fixtures

| Name | What | Size |
|---|---|---:|
| created-empty | `xlsx::Workbook::create()` → `to_bytes` → `from_bytes` | 2,077 B |
| created-dense10k | 10,000 numeric cells (A1:CV100) committed once | 30,012 B |
| MatrixFormulaEval | `test-data/ooxml/xlsx/MatrixFormulaEvalTestData.xlsx`: A1:S138, 783 cells, array formulas, calcChain | 14,875 B |
| trash_item | `test-data/poi/test-data/spreadsheet/Excel_file_with_trash_item.xlsx`: 5,440 cells, an opaque `[trash]/0000.dat` member | 47,318 B |
| no_drawing_patriarch | `test-data/poi/test-data/spreadsheet/no_drawing_patriarch.xlsx`: 75,770 cells, 3.38 MB worksheet XML, 3.44 MB shared strings | 672,414 B |
| gen-10000p | DOCX, 10,000 short paragraphs (`Package::new()` + `add_paragraph_with_text`), main part 850,717 B | 37,290 B |
| drawing.docx, testComment.docx | `test-data/ooxml/docx/` | 95,247 B; 65,298 B |
| gen-200x8 | PPTX, 200 slides × 8 text boxes (`Package::new()` + `add_text_box`) | 229,853 B |

Edits on the real XLSX files go to a 100-column grid below the existing data.
N targets are spread evenly over the document.

## XLSX eager: one transaction (X1) against one transaction per item (X2)

`Workbook::edit` → N × `WorksheetEdit::set` → `commit`; edit start ≤ 1.4 µs
and staging 0.06–0.08 µs per `set` on every base. X2 chains N single-cell
transactions, each from the previous commit's workbook.

| Base | N | X1 edit+set×N+commit | X1 `to_bytes` | X2 chained total | X2 ÷ X1 |
|---|---:|---:|---:|---:|---:|
| created-empty | 1 | 23.4 (22.5–33.5) | 23.8 | 20.9 | 1× |
| created-empty | 10 | 31.5 (30.6–41.8) | 26.9 | 267 | 8× |
| created-empty | 100 | 116 (113–127) | 60.6 | 7,280 | 63× |
| created-empty | 1,000 | 547 (538–570) | 385 | 177,000 | 324× |
| created-empty | 10,000 | 5,230 (5,200–5,430) | 6,000 | — | — |
| created-dense10k | 1 | 1,090 (1,080–1,090) | 6,780 | 1,090 | 1× |
| created-dense10k | 10 | 1,120 (1,110–1,130) | 6,780 | 21,400 | 19× |
| created-dense10k | 100 | 1,250 (1,240–1,280) | 6,190 | 225,000 | 180× |
| created-dense10k | 1,000 | 1,890 (1,870–1,910) | 5,950 | 2,280,000 | 1,206× |
| created-dense10k | 10,000 | 8,900 (8,870–8,970) | 5,960 | — | — |
| MatrixFormulaEval | 1 | 1,860 (1,850–1,880) | 387 | 1,870 | 1× |
| MatrixFormulaEval | 10 | 1,890 (1,870–1,910) | 395 | 19,400 | 10× |
| MatrixFormulaEval | 100 | 2,140 (2,120–2,170) | 454 | 209,000 | 98× |
| MatrixFormulaEval | 1,000 | 3,480 (3,440–3,510) | 814 | 2,770,000 | 797× |
| MatrixFormulaEval | 10,000 | 18,000 (17,900–18,300) | 6,420 | — | — |
| trash_item | 1 | 6,720 (6,690–6,760) | 2,540 | 6,730 | 1× |
| trash_item | 10 | 6,750 (6,710–6,800) | 2,550 | 113,000 | 17× |
| trash_item | 100 | 6,940 (6,920–7,120) | 2,600 | 1,190,000 | 172× |
| trash_item | 1,000 | 7,800 (7,750–7,830) | 2,960 | 12,600,000 | 1,610× |
| trash_item | 10,000 | 16,500 (16,500–17,000) | 8,560 | — | — |
| no_drawing_patriarch | 1 | 53,500 (53,200–55,700) | 41,500 | 53,400 | 1× |
| no_drawing_patriarch | 10 | 53,300 (53,100–54,300) | 41,500 | 1,040,000 | 20× |
| no_drawing_patriarch | 100 | 53,800 (53,300–54,900) | 41,600 | not run | — |
| no_drawing_patriarch | 1,000 | 54,600 (54,300–55,700) | 42,100 | not run | — |
| no_drawing_patriarch | 10,000 | refused (below) | — | — | — |

## XLSX re-open (X3)

`Workbook::from_bytes` of the published X1 N=1 bytes, then the first
one-cell transaction on the fresh workbook (cold) against the same
transaction on an already-used base (warm, X1 N=1).

| Base | `from_bytes` | first transaction, cold | same, warm |
|---|---:|---:|---:|
| created-empty | 32.5 | 24.3 | 23.4 |
| created-dense10k | 150 | 2,190 | 1,090 |
| MatrixFormulaEval | 151 | 2,870 | 1,860 |
| trash_item | 290 | 11,800 | 6,720 |
| no_drawing_patriarch | 3,280 | 89,400 | 53,500 |

## XLSX source-backed value editor (X4)

Open → `try_into_source_backed_package` → `from_source_backed_package` →
`edit` → N × set or insert → `commit` → `publish_commit_to_stream`.

| Base | N | open | edit start | stage | commit | publish | total |
|---|---:|---:|---:|---:|---:|---:|---:|
| created-dense10k, set | 1 | 49.8 | 3,770 | <1 | 3,700 | 6,820 | 14,400 |
| created-dense10k, set | 256 | 53.0 | 3,770 | 54.8 | 4,140 | 7,230 | 15,300 |
| created-empty, insert | 1 | 26.1 | 14.4 | <1 | 13.0 | 33.9 | 87.7 |
| created-empty, insert | 256 | 28.6 | 15.1 | 37.1 | 222 | 114 | 418 |
| MatrixFormulaEval, set (admitted cells) | 1 | 114 | 1,330 | <1 | 1,060 | 518 | 3,020 |
| MatrixFormulaEval, set (admitted cells) | 243 | 116 | 1,330 | 45.9 | 1,350 | 665 | 3,510 |
| MatrixFormulaEval, insert | 1 | 115 | 1,330 | 1.40 | 1,760 | 638 | 3,840 |
| MatrixFormulaEval, insert | 256 | 117 | 1,340 | 180 | 2,260 | 761 | 4,660 |
| no_drawing_patriarch, insert | 1 | 183 | 71,900 | 59.5 | 73,700 | 55,800 | 202,000 |
| no_drawing_patriarch, insert | 256 | 181 | 71,600 | 17,900 | 73,700 | 55,700 | 219,000 |

## DOCX: one transaction (D1)

`edit_document` → N × `replace_paragraph_text` → `commit` → publication, on
gen-10000p. Source-backed publication is `publish_document_commit_to_stream`
into a `Vec`; eager publication is `publish_document_commit` + `to_stream`.
The total excludes the open.

| Owner | N | open | edit start | stage | commit | publish | total |
|---|---:|---:|---:|---:|---:|---:|---:|
| source-backed | 1 | 58.3 | 2,170 | 29.1 | 2,040 | 5,020 | 9,260 |
| source-backed | 10 | 59.7 | 2,170 | 268 | 2,050 | 7,190 | 11,700 |
| source-backed | 100 | 60.5 | 2,190 | 2,650 | 2,090 | 7,110 | 14,000 |
| source-backed | 1,000 | 62.0 | 2,220 | 26,200 | 2,320 | 7,210 | 38,100 |
| source-backed | 4,096 | 64.5 | 2,220 | 105,000 | 3,100 | 7,120 | 118,000 |
| eager | 1 | 210 | 2,110 | 29.8 | 2,090 | 41.4 + 802 | 5,080 |
| eager | 10 | 210 | 2,090 | 270 | 4,100 | 41.7 + 807 | 7,320 |
| eager | 100 | 212 | 2,170 | 2,650 | 4,140 | 43.6 + 809 | 9,850 |
| eager | 1,000 | 216 | 2,160 | 26,200 | 4,320 | 47.3 + 833 | 33,600 |
| eager | 4,096 | 228 | 2,140 | 105,000 | 5,000 | 55.9 + 850 | 113,000 |

## DOCX re-open chain (D2) and opens

Each chained item: source-backed open of the previous output, `edit_document`,
one replacement, commit, publication into a `Vec`.

| N | chain total | per item |
|---:|---:|---:|
| 1 | 9,380 | 9,380 |
| 10 | 93,900 | 9,390 |
| 100 | 943,000 | 9,430 |

| File | source-backed open | then `edit_document` | eager open from bytes |
|---|---:|---:|---:|
| gen-10000p | 35.2 | 2,250 | 176 |
| drawing.docx | 69.4 | 1,120 | 282 |
| testComment.docx | 49.5 | 28.7 | 79.7 |

## PPTX eager (P1)

`Package::from_vec` → `opened_presentation_transaction` → N ×
`set_shape_text` → `commit` → `apply_opened_presentation_commit` →
`to_bytes`, on gen-200x8. "spread" puts one target per slide; "packed" fills
slides in order. The total excludes the open.

| Layout | N | open | edit start (cold) | stage | commit | apply | `to_bytes` | total |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| spread (1 slide) | 1 | 2,250 | 4,820 | 92.7 | 1,520 | 295 | 300 | 7,030 |
| spread (10 slides) | 10 | 2,260 | 4,790 | 732 | 1,820 | 381 | 549 | 8,270 |
| packed (2 slides) | 10 | 2,250 | 4,780 | 733 | 1,540 | 300 | 334 | 7,700 |
| spread (100 slides) | 100 | 2,260 | 4,780 | 6,940 | 4,820 | 1,300 | 2,740 | 20,600 |
| packed (13 slides) | 100 | 2,270 | 4,780 | 6,970 | 1,930 | 414 | 620 | 14,700 |
| spread (125 slides) | 1,000 | 2,280 | 4,790 | 69,200 | 5,660 | 1,550 | 3,380 | 84,600 |

Re-opening the published bytes: eager `from_vec` 2,220, then a cold
transaction start 4,760 (warm, with the digest memo set: 3,570); eager
`from_bytes` (borrowed) 2,670; `SourceBackedPresentation::from_read_at`
1,810, then `try_into_editor` 14.7.

## PPTX source-backed editor

The open (1,830–1,840) is excluded.

| Case | N | edit start | stage | commit | publish | total |
|---|---:|---:|---:|---:|---:|---:|
| one slide, one shape | 1 | 46.9 | 52.5 | <1 | 297 | 397 |
| batch, 1 slide × 8 shapes | 8 | — | 102 | <1 | 305 | 409 |
| batch, 8 slides × 8 shapes | 64 | — | 700 | <1 | 554 | 1,250 |
| batch, 32 slides × 8 shapes | 256 | — | 2,710 | 2.80 | 1,340 | 4,060 |

## Refusals met

1. DOCX, 10,000 operations in one `Edit`, both owners: operation 4,097 →
   `Limit { resource: "operations", max: 4096, actual: 4097 }`
   (`MAX_OPERATIONS`, `crates/litchi-docx/src/document/transaction.rs:47`).
2. XLSX eager, 10,000 cells on no_drawing_patriarch: commit →
   "XML publication rejected for '/xl/worksheets/sheet1.xml': XML Attributes
   limit 250000 exceeded by 250001 at byte 3650391".
3. XLSX source-backed value editor: operation 257 → "value-only batch exceeds
   256 unique cells".
4. XLSX source-backed `set` on MatrixFormulaEval: 119 of 362 numeric cells →
   "the cell belongs to a range-scoped formula".
5. XLSX source-backed on trash_item: every N refused at publication →
   "topology publication refuses non-Part or opaque physical members"; the
   eager route publishes the same file.
6. XLSX survey of 137 corpus workbooks: the source-backed value editor's
   first `edit` is accepted on 45 (79 "value-only edits currently require
   exactly one worksheet", 12 worksheet-relationship refusals, 1 shared-string
   count); the eager `edit` → `commit` → `to_bytes` route is accepted on 132
   (3 protected sheets, 1 invalid dimension '1:5', 1 "owned OPC source
   preservation is unavailable").
7. PPTX source-backed: a second operation on one slide edit →
   `UnsafeEdit { operation: "set_shape_texts", reason: "source-backed slide
   edits support one atomic shape text operation" }`; a 33rd slide in one
   batch → "source-backed slide batch selections exceeds the limit of 32".

## Leaf calls timed (`path:line` in the tree)

- XLSX (`crates/litchi-xlsx/src/`): `Workbook::create` `workbook/model.rs:441`,
  `from_bytes` `:482`, `to_bytes` `:1204`, `edit` `:1361`; `Edit::sheet`
  `workbook/edit/semantic/transaction.rs:595`, `Edit::commit` `:1501`;
  `WorksheetEdit::set` `workbook/edit/semantic/worksheet.rs:352`;
  `Commit::workbook` `workbook/edit/model.rs:2916`; `SourceBackedWorkbook::from_read_at`
  `workbook/source.rs:425`, `try_into_source_backed_package` `:784`;
  `cell_values::SourceBackedEditor::from_source_backed_package`
  `cell_values/source.rs:442`, `edit` `:454`, `publish_commit_to_stream`
  `:537`; `SourceEdit::insert` `:681`, `set` `:686`, `commit` `:828`.
- DOCX (`crates/litchi-docx/src/`): `Package::new` `package/codec.rs:58`,
  `from_reader` `:422`, `to_stream` `:630`; `edit_document`
  `package/package/document.rs:36`, `publish_document_commit` `:227`;
  `source_backed::Package::from_read_at` `source_backed.rs:441`,
  `edit_document` `:1253`, `publish_document_commit_to_stream` `:1619`;
  `Edit::replace_paragraph_text` `document/transaction.rs:2281`,
  `Edit::commit` `:5067`; writer `add_paragraph_with_text`
  `writer/doc/model.rs:450`.
- PPTX (`crates/litchi-pptx/src/`): `Package::new` `package/codec.rs:38`,
  `from_vec` `:225`, `from_bytes` `:246`, `to_bytes` `:360`;
  `opened_presentation_transaction` `package/model.rs:310`,
  `apply_opened_presentation_commit` `:564`; `Transaction::set_shape_text`
  `opened/transaction.rs:217`, `commit` `:1390`;
  `SourceBackedPresentation::from_read_at` `presentation/source.rs:1058`,
  `try_into_editor` `:1250`, `SourceBackedPresentationEditor::from_read_at`
  `:1746`, `edit_slide` `:1913`, `edit_slides` `:1942`,
  `publish_slide_commit_to_stream` `:1956`,
  `publish_slide_batch_commit_to_stream` `:2042`; writer `add_text_box`
  `writer/slide.rs:184`.
