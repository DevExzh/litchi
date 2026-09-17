# 0671: DOC admission residues

Change 0640 removed two field-table refusals and exposed two independent DOC
admission residues. This record closes the row 16 follow-up assigned by change
0651 and decision 1 of change 0652. It makes one narrow correctness change for
the `PapxInFkp` witness, keeps the `FBKF.ibkl` uniqueness check strict after
reading the primary specification, and exposes the existing stylesheet
leniency option through the unified `litchi::Document` facade.

Disposition: retained, correctness fix. `performance_claim: none`.

## `FBKF.ibkl`: the uniqueness refusal stays

`ole/doc/watermark.doc` has nine `FBKF` records with `ibkl` values
`[0, 1, 4, 4, 2, 3, 6, 7, 8]`. The repeated value is in range; the two
referenced `PlcfBkl` entries both contain CP 535, and the unused slot also
contains CP 535. The bytes therefore describe a determinable bookmark table,
but that coincidence does not remove the format's uniqueness requirement.

The current Microsoft `[MS-DOC]` definition of [`FBKF`](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/4dfad7b0-37bb-443a-8933-3bb79d2c5994)
says that `ibkl` is a zero-based index into `PlcfBkl` and **MUST be unique for
all FBKFs inside a given `Plcfbkf` or `Plcfbkfd`**. The companion
[`Plcfbkf`](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/82702014-0081-40af-8d4e-421db135f3ff)
description gives the one-to-one FBKF/PlcfBkl relationship and permits equal
bookmark CPs for overlapping bookmarks. Equal CPs explain why the witness is
recoverable, but they do not satisfy the explicit `ibkl` MUST.

The `HashSet` guard in `crates/litchi-doc/src/parts/bookmarks.rs` is therefore
unchanged. The facade regression test keeps the refusal typed and checks the
existing message `bookmark ibkl values must be unique and in range`. This is a
resolved specification question, rather than a new leniency case.

## `PapxInFkp`: accept one proven word-alignment byte at the PAPX boundary

The POI witness has `cb = 0`, `cb' = 3`, and the six bytes
`00 00 | 31 24 00 00`: `istd = 0`, one complete three-byte SPRM, and a final
zero. Every one of its 24 PAPX entries has the same final-byte shape. The
[`PapxInFkp`](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/580510b8-df7a-467e-a51c-0d71eb15c7cd)
encoding requires the `cb = 0` form to carry `2 * cb'` bytes, while a nonzero
`cb` carries `2 * cb - 1` bytes. The companion
[`GrpPrlAndIstd`](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/bd96f2aa-1318-4066-9723-4db035ef412b)
grammar requires the grpprl portion to contain whole Prl elements. The
published MS-DOC prose does not separately name this final alignment byte, so
the admission is grounded in the repeated witness shape and the established
Word-compatible writer behavior documented by [wv2's PAPX generator](https://fossies.org/linux/wv2/src/generator/generator_wword8.htm),
which pads an odd grpprl with one zero byte to a word boundary.

`PapBinTable::parse` now applies `trim_papx_word_alignment_pad` only to the
`GrpPrlAndIstd` bytes obtained from a PAPX FKP. It removes one byte only when
all of these facts hold: the stored sequence is even-sized, ends in zero,
strict `parse_sprms` reports exactly one remaining opcode byte at the end, and
the preceding bytes are therefore a complete SPRM sequence. Nonzero `cb`
encodings, valid sequences ending in a zero operand, truncated operands,
invalid lengths, nonzero tails, indirections and cycles retain the existing
typed failure. The shared `parse_sprms` function remains exact; a direct call
with the same trailing zero still returns `Error::Opcode`.

This is a compatibility admission at the format boundary, not a general
repair rule. The fixture now opens through the unified facade and returns
nonempty text. No writer change is made, so authored PAPX output continues to
use its existing exact length encoding.

## Facade access to stylesheet leniency

`litchi::Document` now provides `open_with_doc_options` and
`from_bytes_with_doc_options` behind the existing `doc` feature. Both accept
the already public `litchi::doc::OpenOptions`; callers can select
`litchi::doc::Leniency::TolerateStylesheetDefects` without bypassing the
unified facade. The path method uses the same bounded filesystem fallback as
the ordinary document path. Options are consumed by the DOC reader only;
other detected formats continue through their normal facade route, and the
default `open`/`from_bytes` behavior is unchanged.

The duplicate-style fixture remains rejected by the default facade and opens
through both new methods with the existing leniency report behavior. Structural
DOC errors remain fatal under the leniency option.

## Validation and limits

The worktree contains 57 `.doc` fixtures, matching the 0640 corpus inventory.
The retained 0640 structural scan remains the corpus-level evidence for the
two original witnesses (42 admitted and 15 refused on both 0640 legs); this
follow-up adds direct tests for both residues and does not claim a new full
corpus differential or any performance result.

Validation on this branch:

* `cargo fmt --all --check`;
* `CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_BUILD_JOBS=2 cargo test -p litchi-doc --lib` — 1,007 passed, 2 ignored;
* the `litchi-doc` `doc_leniency` integration test — 4 passed;
* `CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_BUILD_JOBS=2 cargo test -p litchi --features doc --lib` — 31 passed;
* `CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_BUILD_JOBS=2 cargo check -p litchi --features doc,docx` — passed.

The remaining specification gap is explicit: the current primary MS-DOC pages
define the two PAPX length forms and require whole Prl elements, but do not
give a standalone normative sentence for an alignment pad. The narrow rule is
kept because the witness is repeated across all PAPX entries and compatible
implementations document the same pad shape; a future primary clarification
could refine the comment without widening the parser. No `litchi-cfb` path or
limit was changed.
