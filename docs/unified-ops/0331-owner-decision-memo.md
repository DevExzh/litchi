# 0331 — Owner decision memo (2026-10-17)

Status: decided by the owner on 2026-10-20 (record
[0350](0350-owner-api-debt-decision.md)): items 14–20, 23, 24, 26 and 29
are overridden (required work); item 22 is deferred; every other item is
confirmed. Originally: awaiting owner confirmation. Every item below is already in effect
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

## Second-audit additions (record 0344)

Source: the second definition-of-done audit (2026-10-18, §12.2 items 3
and 4). Items 14–24 are record 0343's (the API deferrals D1–D11); these
continue the numbering after them. Each is in effect as the adopted
default, like the items above.

| # | Decision | Adopted default | Record |
|---|---|---|---|
| 25 | XLS and ODS element parity (audit WP10/WP11): implement it in v1, or defer it | Deferred. The 392 P1 capability-delta rows still `planned` (XLS 196, ODS 196: chart, chart series and axes, run, table and detected-table families with their properties) are `deferred-P1`, because both formats are outside the v1 boundary (`policy.p0Formats` = docx, xlsx, pptx). Their follow-ups name records 0327 and 0316 (XLS) and 0327 (ODS). No P0 row changed: the delta keeps 26 P0 rows `planned` (DOCX, PPTX) and 4 out-of-scope | 0344, 0327, 0316 |
| 26 | `litchi::sheet::text` against design §16.4 ("re-homed behind `import` semantics in M4") | Kept, as documented non-Office tooling: its CSV, TSV, PRN, SYLK, DIF and fixed-width readers and writers open none of the 21 formats, and `File::import` (record 0266) covers only CSV/TSV into a workbook; deleting it drops those capabilities and deprecation is not allowed. The follow-up is to move it into a leaf crate and align its names with ADR 0034 | 0334 §2 item 2 |
| 27 | The property-level gap enumeration: does every OfficeCLI property Litchi does not read need a `capabilities()` row? | No. The generated capability delta (`officecli-capability-delta.md` and its per-format pages) is the property-level gap list: its 3,216 `unlisted` property-read rows (5 P0: DOCX `comment.anchoredTo`; PPTX `modernComment.initials`/`resolved`/`parent`/`id`; 62 P2, the rest P1) each carry a disposition, and `--check` fails on a row without one. `capabilities()` lists elements, operations and the property gaps the registry knows; it does not mirror OfficeCLI's vocabulary | 0332, 0339, 0344 |
| 28 | The bench misses explained as residuals | Accepted. `results/0289/bench-matrix.md` lists 144 open misses (trunk after record 0346), each naming a follow-up or an explained residual (the 54 of record 0337, inherent to the comparison or kept by an earlier owner decision, plus record 0338's correctly rounded JSON number parse and record 0329's Keynote layout noise). Under 0248 D5 (item 1) a residual stands only with this acceptance; a residual the owner rejects becomes a follow-up | 0337, 0344 |
| 29 | `Box<dyn>` in the `litchi-core::sheet` traits (goal constraint 2: no `Box<dyn Trait>` on a public or per-operation path) | A recorded exception: `litchi-ops` implements the pre-existing object-safe `litchi-core::sheet` traits, which return `Box<dyn CoreWorksheet>`/`Box<dyn CoreCell>` (`calc.rs`, `view/workbook.rs`, `format/{xlsx,xlsb,xls}`; 72 occurrences), so the evaluate path builds one box per cell. Replacing them with static dispatch is a separate change (size L) | audit §1 item 8c |
| 30 | DOCX chart series and axis formatting breadth (record 0335 F4) | Deferred-P1. The 34 capability-delta rows for series formatting and axis formatting keys beyond name/values/color and title/min/max were `planned`; they are P1 formatting breadth under 0332's rule (the series and axis elements themselves are supported). With this, no P0 or P1 row remains `planned` | 0335, coordinator 2026-10-19 |

Of these, 25 and 27 decide whether the parity layer of design §14 counts
as met; 29 is the only one that touches a goal constraint.
