# Log sections for change 0671

Four paragraphs for the coordinator to merge into the shared program logs.
Change 0671 is a correctness fix; `performance_claim: none`.

## For `HOTSPOTS.md`

## 0671 — one DOC residue was a real MUST and the other was a bounded compatibility byte

Record: [0671](../../0671-doc-admission-residues.md). Change 0640's two
remaining DOC refusals now have different dispositions. `watermark.doc` repeats
`FBKF.ibkl` value 4 while leaving slot 5 unused; both `PlcfBkl` slots contain CP
535, so the table is recoverable, but Microsoft's current [MS-DOC `FBKF`]
definition says `ibkl` **MUST be unique** within a `Plcfbkf` or `Plcfbkfd` and
the strict typed refusal stays. The POI `test.doc` witness has `cb=0`, `cb'=3`,
an `istd`, one complete odd SPRM and a final zero in every one of its 24 PAPX
entries. `PapBinTable` now drops one final zero only for an even `GrpPrlAndIstd`
whose strict SPRM parse proves that byte is exactly one incomplete opcode; all
other malformed tails remain errors and `parse_sprms` is unchanged. The witness
remains a typed refusal on the strict default route and opens with the explicit
`OpenOptions::with_papx_alignment_padding()` profile. The unified facade adds
bounded `Document::open_with_doc_options` and `from_bytes_with_doc_options`,
making the existing stylesheet leniency and PAPX compatibility profile opt-in
reachable without changing strict defaults.
The checked-in MS-DOC snapshot confirms the FBKF rule at §2.9.70 (local line
6648), but its §2.9.175 PapxInFkp and §2.9.114 GrpPrlAndIstd text (local lines
18684-18688 and 12469-12471) requires whole Prl elements and does not specify
this PAPX pad; the allowance is therefore compatibility evidence, not a new
normative claim. No timing or performance result is claimed. [Record and
limitations](../../0671-doc-admission-residues.md).

[MS-DOC `FBKF`]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/4dfad7b0-37bb-443a-8933-3bb79d2c5994

## For `GOAL_AUDIT.md`

## 0671 — strict malformed-input handling is preserved while one format-boundary refusal moves

Record: [0671](../../0671-doc-admission-residues.md). `docs/GOAL.md` requires
correctness, lossless semantics and typed fail-closed behavior before speed.
The FBKF check follows the primary specification rather than the witness's
recoverability: equal CP values do not legalize a repeated index, so the
duplicate remains unavailable under strict reading. The PAPX change is bounded
to one consumer and one shape, and is enabled only by
`OpenOptions::with_papx_alignment_padding()`. It sees an even `cb=0` payload
ending in zero, asks the exact shared SPRM parser whether the prefix is complete
with one byte left, and trims only that byte; a malformed operand, length,
nonzero tail, indirection or cycle still reaches the existing typed error. The
global SPRM parser remains strict, and the writer remains unchanged. The facade
carries the existing `OpenOptions` and `Leniency` type through both path and byte
entry points, while structural defects remain fatal under stylesheet tolerance.
checked-in MS-DOC snapshot's §2.9.175 (local lines 18595 and 18684-18688)
defines the PAPX length forms and §2.9.114 (local lines 12380 and 12469-12471)
requires whole-Prl grammar, while neither names this alignment byte. The record
leaves that specification gap visible and relies on repeated fixture bytes plus
compatible implementation evidence behind an explicit format-owned option; the
strict default still refuses the witness. `performance_claim: none`.

## For `REPORT.md`

## 0671 — the DOC fixtures now stop at the right boundary

Record: [0671](../../0671-doc-admission-residues.md). The 0640 packet left two
DOC refusals provisionally unresolved. The first is still a refusal: the
watermark's repeated `ibkl` maps two FBKFs to the same `PlcfBkl` CP, but the
current Microsoft definition makes that index unique. The second was an
alignment boundary: POI's `test.doc` stores six bytes for `istd` plus a
three-byte paragraph SPRM, with one zero completing the `cb=0` even-byte
container; all 24 PAPX entries repeat it. The strict consumer still refuses
that shape by default; the explicit `OpenOptions::with_papx_alignment_padding()`
profile accepts precisely it, and direct tests prove a nonzero trailing byte
and an uncontextualized trailing zero still fail. The unified facade's new
options methods admit the existing duplicate-style fixture only when the caller
opts into `TolerateStylesheetDefects`; default facade reads remain strict.
Validation is
`litchi-doc` 1,007 passed/2 ignored, its four leniency integration tests passed,
the DOC facade's 31 unit tests passed, and the combined `doc,docx` facade check
passed. The repository's 57-file DOC inventory and 0640 corpus scan are retained
as the corpus baseline; this follow-up does not claim a new full differential,
timing, instruction, allocation or RSS number. The checked-in MS-DOC reference
confirms FBKF uniqueness at §2.9.70 (local line 6648) and the PAPX length/whole-
Prl constraints at §§2.9.175 and 2.9.114 (local lines 18684-18688 and
12469-12471); it does not make the alignment-byte allowance normative.

## For `ADR_COMPLIANCE.md`

## 0671 — ADR 0003 and ADR 0006 applied at the format boundary

Record: [0671](../../0671-doc-admission-residues.md). The implementation follows
ADR 0003's typed refusal rule and ADR 0006's reader/writer division. The reader
keeps the normative `FBKF.ibkl` uniqueness refusal and does not convert it into
leniency. For PAPX, the reader accepts one producer alignment byte only after a
strict parser proves the preceding sequence complete and only under the explicit
`OpenOptions::with_papx_alignment_padding()` profile; the default remains a
typed refusal. No writer emits or normalizes that byte, and `parse_sprms` retains
its exact-sequence contract. The new facade methods expose the public DOC
options and preserve the strict default, with the path read using the existing
bounded fallback helper. No CFB entry point, resource limit, error variant or
archive state is changed. The
primary local MS-DOC sections make the uniqueness rule explicit (§2.9.70, local
line 6648) and define PAPX length arithmetic (§2.9.175, local lines
18684-18688), while the whole-Prl requirement (§2.9.114, local lines
12469-12471) means the observed alignment-byte behavior remains a documented
compatibility inference rather than a new normative claim. The record therefore
names the open specification question and the corpus boundary instead of treating
the one witness as universal. `performance_claim: none`.
