# 0331 — Owner decision memo (2026-10-17)

Status: awaiting owner confirmation. Every item below is already in effect
as the adopted default, so no work waits on it. The owner can confirm an
item, or override it; an override becomes a numbered record and follow-up
work.

Source: the coordinator's definition-of-done audit (2026-10-17, §7), the
"Awaiting owner confirmation" and "Awaiting owner decision" sections of
`PROGRESS.md`, and the records cited.

| # | Decision | Adopted default | Record |
|---|---|---|---|
| 1 | Tier strictness, provisional decisions, fix-before-explain | D3 tiers stay strict; D4 earlier provisional decisions stay provisional; D5 attempt a fix before writing an explanation | 0248 |
| 2 | Removing a DOCX `num` | The paragraphs' `numPr` references to it are kept. OfficeCLI clears them, which edits paragraphs the caller did not name | 0319 |
| 3 | Re-adding a built-in DOCX style id | Refused, typed (corpus delta D11). OfficeCLI upserts | 0319 |
| 4 | XLS CONDFMT rule count > 3 and out-of-place ShrFmla | Kept spec-strict (refused) | 0316 |
| 5 | M0a S10 F1/F2 | Explanations E1 (first-call harness rows) and E2 (RTF open allocator state) accepted | s10-measurement §9 |
| 6 | M2 kernel defaults D1–D19, Q1–Q14 | Verbs take `&self` and return a new `File`; strict refusal by default; save after `unsign` writes the unsigned copy; XLSX promotes on first write; raw-layer timing per Q14 | 0015 |
| 7 | Regex and sheet-name matching | Q3: regex values are refused until the `regex` crate is approved. Q6: sheet names match case-insensitively in every workbook format. The other eleven §11.3 defaults are as recorded | 0010 |
| 8 | Shared handles | A shared handle shares the leaf package; there are no consuming verbs | 0009 R4 |
| 9 | Early defaults | F2 flat `File`, Q1 `limit_exceeded`/`cancelled`, Q3 `retain_source = true`, Q4 `Signed` wraps signed input. **Q2 (iWork → `unsupported_format`) is withdrawn**, as owner decision D1 (2026-10-15) superseded it | 0007 |
| 10 | Surface shape and names | One `OpenOptions` with per-format sub-options; verb `apply`; find/replace through props; `merge` means template merge | design, 0015 |
| 11 | Core `Limits` | The glossary ban on the core `Limits` name stands | 0249 |
| 12 | XLSX SmartArt tier | Deferred to P1 as a typed capability gap. Design §13 lists typed SmartArt CRUD as "also in v1" because the models exist, but `litchi-xlsx` has no diagram model, and no OfficeCLI corpus line uses XLSX SmartArt | 0339 |
| 13 | Picture and table breadth | PPTX picture crop/effects, `table.style` and DOCX picture crop/floating/link are deferred-P1 under 0332's breadth rule (a P0 element's properties are P1 unless design §13 names them); the picture elements stay P0 and supported | 0332, 0333, 0339 |
| 14 | Direct `&mut self` package mutators (DOCX 85, PPTX 77, XLSB `into_workbook` and mutating `Workbook` verbs) | Kept; no new public `&mut self` mutator may be added to a `Package` | 0343 D1 |
| 15 | Scoped OOXML commit triad (`snapshot()` on ~200 scoped commits, missing `Diagnostics`, DOCX/XLSX scoped `Transaction` names) | Kept until each family is migrated to publish a `Package`; the primary transactions are canonical | 0343 D2 |
| 16 | PPT scoped `Transaction` families and `rollback` | Kept as W4 residue | 0343 D3 |
| 17 | DOCX/PPTX `Error::UnsafeEdit` refusal vocabulary; ODT `MutableDocument` | Kept; the replacement error taxonomy is an owner choice | 0343 D4 |
| 18 | PPT `raw::Reader` and 173 `*_with_limits` entry points | Kept; not folded into `OpenOptions` | 0343 D5 |
| 19 | Flat ODF owners' limits (`FlatDrawing`/`FlatPresentation` fixed ceilings; crate `Limits`, not `OpenOptions`) and ODC chart `Limits` | Fixed ceilings and crate `Limits` stay | 0343 D6 |
| 20 | `Metadata.keywords` type | `Option<String>` stays, because the type fixes the JSON/CLI/Python metadata shape; the OOXML codec keeps keywords losslessly | 0343 D7 |
| 21 | DOCX app/custom properties through the shared OOXML engine | Withdrawn: implemented by record 0342 (`metadata::Edit::{set_app_props, set_custom_props}`, `Package::app_props()`) | 0343 D8, 0342 |
| 22 | Content-redaction audit of OOXML/legacy error strings | The rule stands; violations are fixed as found | 0343 D9 |
| 23 | iWork `CellPosition`/`CellRange` → `At`/`Area` | Deferred to the next iWork wave (unblocked: record 0340 has merged) | 0343 D10 |
| 24 | `litchi-iwa` `*DocumentBuilder` prefixes | Kept; rename or delete the unpublished host crate is the owner's call | 0343 D11 |

The questions most likely to change behaviour if overridden are 2, 3, 4,
7 (Q3, regex) and 12; of the 0343 API deferrals, 14, 15 and 20 have the
widest public-API reach.
