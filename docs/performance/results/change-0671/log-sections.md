# Log sections for change 0671

Four paragraphs for the coordinator to merge into the shared program logs.
Change 0671 is a correctness fix; `performance_claim: none`.

## For `HOTSPOTS.md`

## 0671 — one DOC residue was a real MUST and the other was a word-boundary byte

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
opens and returns text. The unified facade adds bounded
`Document::open_with_doc_options` and `from_bytes_with_doc_options`, making the
existing stylesheet leniency opt-in reachable without changing strict defaults.
No timing or performance result is claimed. [Record and limitations](../../0671-doc-admission-residues.md).

[MS-DOC `FBKF`]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/4dfad7b0-37bb-443a-8933-3bb79d2c5994

## For `GOAL_AUDIT.md`

## 0671 — strict malformed-input handling is preserved while one format-boundary refusal moves

Record: [0671](../../0671-doc-admission-residues.md). `docs/GOAL.md` requires
correctness, lossless semantics and typed fail-closed behavior before speed.
The FBKF check follows the primary specification rather than the witness's
recoverability: equal CP values do not legalize a repeated index, so the
duplicate remains unavailable under strict reading. The PAPX change is bounded
to one consumer and one shape. It sees an even `cb=0` payload ending in zero,
asks the exact shared SPRM parser whether the prefix is complete with one byte
left, and trims only that byte; a malformed operand, length, nonzero tail,
indirection or cycle still reaches the existing typed error. The global SPRM
parser remains strict, and the writer remains unchanged. The facade carries the
existing `OpenOptions` and `Leniency` type through both path and byte entry
points, while structural defects remain fatal under stylesheet tolerance. The
current MS-DOC pages define the PAPX length forms and whole-Prl grammar but do
not explicitly name this alignment byte; the record leaves that specification
gap visible and relies on repeated fixture bytes plus compatible implementation
evidence for this narrow admission. `performance_claim: none`.

## For `REPORT.md`

## 0671 — the DOC fixtures now stop at the right boundary

Record: [0671](../../0671-doc-admission-residues.md). The 0640 packet left two
DOC refusals provisionally unresolved. The first is still a refusal: the
watermark's repeated `ibkl` maps two FBKFs to the same `PlcfBkl` CP, but the
current Microsoft definition makes that index unique. The second was an
alignment boundary: POI's `test.doc` stores six bytes for `istd` plus a
three-byte paragraph SPRM, with one zero completing the `cb=0` even-byte
container; all 24 PAPX entries repeat it. The consumer now accepts precisely
that shape, and direct tests prove a nonzero trailing byte and an uncontextualized
trailing zero still fail. The unified facade's new options methods admit the
existing duplicate-style fixture only when the caller opts into
`TolerateStylesheetDefects`; default facade reads remain strict. Validation is
`litchi-doc` 1,007 passed/2 ignored, its four leniency integration tests passed,
the DOC facade's 31 unit tests passed, and the combined `doc,docx` facade check
passed. The repository's 57-file DOC inventory and 0640 corpus scan are retained
as the corpus baseline; this follow-up does not claim a new full differential,
timing, instruction, allocation or RSS number.

## For `ADR_COMPLIANCE.md`

## 0671 — ADR 0003 and ADR 0006 applied at the format boundary

Record: [0671](../../0671-doc-admission-residues.md). The implementation follows
ADR 0003's typed refusal rule and ADR 0006's reader/writer division. The reader
keeps the normative `FBKF.ibkl` uniqueness refusal and does not convert it into
leniency. For PAPX, the reader accepts one producer alignment byte only after a
strict parser proves the preceding sequence complete; no writer emits or
normalizes that byte, and `parse_sprms` retains its exact-sequence contract.
The new facade methods expose an existing public option and preserve the strict
default, with the path read using the existing bounded fallback helper. No CFB
entry point, resource limit, error variant or archive state is changed. The
primary MS-DOC pages make the uniqueness rule explicit and define the PAPX
length arithmetic, while the observed alignment-byte behavior remains a
documented compatibility inference rather than a new normative claim. The
record therefore names the open specification question and the corpus boundary
instead of treating the one witness as universal. `performance_claim: none`.
