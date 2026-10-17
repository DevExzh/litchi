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

The questions most likely to change behaviour if overridden are 2, 3, 4
and 7 (Q3, regex).
