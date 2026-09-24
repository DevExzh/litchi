# 0768: DOC protection is classified from the fields MS-DOC gives protection meaning, a saved caret no longer blocks an insertion at its CP, and FKP pages filled to their size estimate no longer corrupt their BX array: all 35 readable DOC fixtures now classify as unprotected and accept a tracked insertion at CP 0, and every real protection signal is still refused

Status: retained, implemented. `performance_claim: none`. This is a correctness
change (spec conformance, [MS-DOC]), not a safety relaxation: `authorize()`
is unchanged, every document or range protection state is still refused by
default, and every malformed or incomplete host shape is still refused under
every policy. Its cost is measured below as evidence, not registered as a claim.
An independent review found the protection boundary holds and returned
merge-after-fixes; its follow-ups are in section "Independent review and
follow-ups".

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

Base `1d1044e3ac` (after the spec-gap merge, record 0759); branch
`perf/0768-doc-protection-classification`. Commits:

| commit | what it does |
| --- | --- |
| `a4767ba8a5` | `fix(doc)`: classify DOC protection from the MS-DOC protection fields (C1–C3 below); real fixtures replace the `with_valid_word97_dop` normalization in two tests |
| `6b340fe0e9` | `fix(doc)`: keep a saved `Selsf` CP in place for an insertion at that CP; a CP strictly inside replaced or removed text stays refused |
| `ac63852fc3` | `fix(doc)`: stop CHPX and PAPX FKP pages filled to the size estimate overwriting their BX array (found while testing the first two) |
| `e4022968db` | this record, its evidence packet and the two `FEATURE_MATRIX.md` rows the change makes stale |
| `ab45d1d416` | `docs(doc)`: `EditProtection::Unrecognized` also names a DOP outside the table stream |
| `95b485a290` | `fix(doc)`, review (required): a selected inline picture or shape moves after text inserted at its first CP |
| `de2df97e1f` | `fix(doc)`, review: a tracked insertion after the main story's final paragraph mark is refused |
| `0792812830` | `fix(doc)`, review: the public FKP builders refuse a property no page can hold instead of overlapping or panicking |
| `83fa8563d7` | `fix(doc)`, review: the positional body-text source classifies protection with the shared classifier |
| review record commit | the follow-ups' sections here, the second gate run and the packet updates |

Evidence: [results/change-0768](results/change-0768/README.md).

## Result

**Classification.** The strict classifier merged in 0759 (incoming `581711a4c4`,
kept by 0759's merge judgement 8) refused 24 of the 38 checked-in DOC fixtures as
`Unknown` or `Unrecognized`, which no `AllowProtected` capability can bypass.
None of the 24 is protected: in every one `fLockAtn`, `fProtEnabled`, `fLockRev`
and `lKeyProtDoc` are 0, `fEnforceDocProt` is 0 wherever the DOP reaches it, and
there are no range-permission tables and no encryption. The classifier refused
them for grammar that MS-DOC either does not require of a reader or places on
fields with no protection meaning. After this change all 35 readable fixtures
classify as `None`; the three unreadable ones (encrypted, Word 6, invalid CFB)
stay `Unknown`. The per-fixture table is below.

**Tracked insertion at CP 0.** Before: 9 of the 38 fixtures accepted it (11 at
CP 1). After: all 35 readable fixtures do, at CP 0 and CP 1; each output reopens
in the strict editor with the revision present and republishes byte-identically,
and the public reader accepts the output exactly when it accepts the source.

**Wider corpus.** Across all 57 `.doc` files in `test-data`, 17 classified as
`None` before and 49 after; no readable file carries a protection signal.
Tracked insertions now succeed on 12 of the 19 files outside `test-data/ole/doc`
(4 before). Every remaining refusal is protection-independent: encryption, a
pre-Word 97 FIB, an invalid CFB, a malformed SPRM sequence, and one file whose
`Selsf` points beyond its text (an earlier litchi output).

**Adversarial finding: FKP page overflow.** Two files the old classifier refused
(`saved-by-table.doc`, `bookmark-delete-redline.doc`) then produced output the
editor could not reopen ("malformed CHPX FKP"); a third
(`International-Travel-Approval-Request-Form.doc`) already did so at the base.
The CHPX page builder (`writer/fkp.rs`, shared by the fresh writer, the
tracked-revision editor and embedded-object commits) and both PAPX page
builders could place their lowest property one byte into the BX array. Fixed:
all three outputs now reopen, and page layouts that were already valid are
unchanged.

**Cost.** Exact timed-region instruction counts (callgrind) change by −0.83% to
+0.27% on the four control workloads; `classify` itself executes 81% fewer
instructions. Wall clock is within noise on NoHeadFoot.doc and
FloatingPictures.doc. On the harness `large` DOC shape it swings from −20% to
+31% depending on glibc's heap state, and with glibc's malloc thresholds pinned
it is unchanged (median ratio 1.000, page faults equal). Every flag over 5% is
listed under "Regression flags".

**After the review.** A saved inline-picture or shape selection now moves with
its object when text is inserted at it (the first version kept its CP, so the
selection covered the new text); the positional body-text source gives the same
protection verdict as the shared classifier on all 57 DOC files and on the
review's 1,730 crafted records wherever it reads protection; the public FKP
builders refuse a property no page can hold, which ends a panic the fresh writer
reached with a table of 23 or more columns; and a tracked insertion after the
final paragraph mark is refused. The follow-ups were not re-measured.

## Why this record exists

The spec-gap branch (merged in 0759) replaced the pre-merge protection check with
a classifier that parsed the whole typed DOP and required exact FIB and DOP
shapes per generation, with no design record; its evidence packet
(`docs/report/spec-gap-validation-evidence/doc-protection-v5`) lists as
"normative fixes" that nFib 0x0101 requires cswNew 2 and a 594-byte Dop2002,
with "no zero-cswNew/500-byte compatibility exception". The merge kept it and
switched the edit tests to synthetic `with_valid_word97_dop` fixtures. A
read-only analysis for the coordinator then found that the 24 refusals come from
checks stricter than MS-DOC, in fields that carry no protection data, and that 11
further CP-0 refusals come from a `Selsf` check that exists only on the incoming
branch. The coordinator dispatched this record to implement the analysis's
recommendation, refining it where the spec gives better evidence, and to decide
the `Selsf` question against MS-DOC.

## What changed

Scope: `crates/litchi-doc` only. No `unsafe`, no dependency, no limit relaxed,
no public signature changed (`EditProtection::Unrecognized` keeps its meaning;
its documentation now lists what produces it). The review follow-ups add typed
refusals to `RevisionEditor::add_text` and to the public FKP builders, and let
the positional body-text source open a protection shape it cannot prove as
read-only instead of refusing it.

### C1: the FIB (`parts/protection/policy.rs`, `validate_fib_shape`)

- When `cswNew` is 0, the effective `nFib` is `FibBase.nFib` (2.5.14), with
  0x00C0 and 0x00C2 read as 0x00C1 (Appendix A note <11>: the shell's empty
  document and the BiDi build of Word 97).
- A zero `cswNew` is accepted for every generation. 2.5.1 asks for 2 or 5 from
  Word 2000 on, but LibreOffice writes `nFib` 0x0101 with a zero `cswNew`
  (`3rdparty/libreoffice-core/sw/source/filter/ww8/ww8scan.cxx:6286-6292`, and
  :6622-6624 places its `FibRgCswNew` at FIB offset 0x3FA, inside the counted
  pointer array); `FibRgCswNew` carries no protection data; and 2.5.15 has the
  reader take what the counts declare.
- Unchanged: `csw` 0x000E and `cslw` 0x0016; a nonzero `cswNew` must equal the
  generation's count and its `nFibNew` must be present; `cbRgFcLcb` must equal the
  count 2.5.1 assigns to the effective generation (it bounds the pointer array
  that locates the DOP and the range tables, indexes 31 and 141–144); an unknown
  generation is `Unknown`.

### C2: the DOP length (`dop_holds_generation_protection_fields`)

The exact per-generation lengths (500/544/594/616) become a presence rule:

| effective nFib | required `lcbDop` | why |
| --- | --- | --- |
| 0x00C1, 0x00D9, 0x0101 | ≥ 84 | `DopBase` holds these generations' protection fields |
| 0x010C | ≥ 600 | Dop2003 defines the enforcement unit at 598..600; a truncated Dop2003 (595..599) cannot be proven unprotected |
| 0x0112 | 674, 690 or 694 | the one exact length rule MS-DOC states (2.7.1) |

A field beyond the recorded length takes its specified default, which for every
protection field is unrestricted (`fEnforceDocProt` "By default ... 0").

### C3: the DOP content (`dop_restricts_editing`)

The classifier no longer runs `DocumentProperties::parse_bytes` and
`versioned()`. It reads only:

- `DopBase` (2.7.2): `fLockAtn`, `fProtEnabled`, `fLockRev`, and `lKeyProtDoc`;
  any of them set is `Document`.
- `Dop2003` byte 598 (2.7.7), whenever the DOP reaches it, in any generation:
  `fEnforceDocProt` (bit 3) with `iDocProtCur` (bits 4..6) 0–3 is `Document`,
  7 is no restriction. The analysis proposed reading the 16-bit word at 598..600
  whenever `lcbDop` ≥ 600; both fields lie wholly in byte 598, so this reads
  them whenever the DOP is at least 599 bytes, which also covers a crafted
  599-byte DOP that the ≥ 600 rule would have read as unprotected.

The MUSTs MS-DOC places on those fields stay enforced and give `Unrecognized`:
`fLockAtn` with `fLockRev`, `fLockRev` without `fRevMarking`, `fFormNoFields`
without `fProtEnabled` (2.7.2), and a reserved `iDocProtCur` 4..6 whether or
not it is enforced (2.7.7; the base's typed parser refused it either way).
`fProtEnabled` together with `fLockAtn` or `fLockRev` is only a SHOULD NOT that
Word 97–2003 is documented to break (notes <164>, <165>, <167>): it is now
`Document` (refused by default, open to `AllowProtected`) instead of
`Unrecognized`.

No longer validated by the classifier, because they restrict nothing:
`wvkoSaved` (LibreOffice writes 7, `ww8scan.cxx:8169-8170`), `wSpare2` and
`reserved2` ("MUST be 0, and MUST be ignored"), `fpc`/`rncFtn`/`epc`, the DTTMs,
`pctWwdSaved`, the Dop97 `Dogrid` display multiples (Word writes 0), Dop2003
`empty2`, and the Dop2010 `docid` (Word writes 0). The typed DOP API
(`Document::document_properties()`) keeps all of its checks.

### What stays refused (unchanged)

`ProtectionPolicy::authorize`; the range-protection tables (`Ranges::parse`,
malformed → `Unrecognized`); `Unknown` for a truncated or overflowing pointer
array, a wrong `csw`/`cslw`, an unknown generation, a wrong or truncated nonzero
`cswNew`, a `cbRgFcLcb` mismatch, a missing DOP pointer, and `lcbDop` 0; and
`Unrecognized` for a DOP outside the table stream (the base's mapping; both are
non-bypassable).

### The `Selsf` check (`parts/saved_selection.rs`, `remap_splice_cp`)

Determined against MS-DOC 2.9.244: the check refused more than the spec makes
ambiguous. `Selsf` "specifies the last selection that was made to the document";
MS-DOC defines no remapping for it, only field requirements and what each kind of
selection is. When text is inserted at a CP the record holds, the recorded
position stays between the same original characters whichever side of the new
text it takes, so which side is right depends on what the record says the
selection is. (Corrected after the review: the first version of this record
said that keeping the CP was the one uniform choice for every kind.)

- **Keep the CP (the new text follows it):** a character selection, an
  insertion point, a text frame, whole table rows and a text block. A whole-row
  selection's `cpFirst` MUST be the start of its row and its `cpLim` the end of
  its last row, and a text block's `cpFirst` and `cpLim` MUST be line starts;
  text inserted there leaves each of those positions where it was. Tracked text
  has no paragraph mark, so text inserted at a frame's first CP joins the frame's
  first paragraph and the frame still starts there. A caret stays on its
  recorded CP (the `fInsEnd` line-end affinity is kept), and text inserted at a
  character selection's first CP joins it, as a replacement's text already did.
  This is also the rule the remap already applied to the start of a replaced
  range, so a pure insertion is the empty-range case of the same function (the
  special case was deleted).
- **Move after the text (`95b485a290`):** an inline picture (`fGraphics`) is its
  0x0001 character and a shape or floating picture (`fShape`) its 0x0008 anchor
  (1.3.5, 2.8.27). Text inserted at the object's first CP is written in front of
  that character, and the editor moves the `PlcfSpa` anchor with it, so every CP
  of the record at the insertion point moves after the text and the selection
  covers the object and none of the new text. Keeping the CP there, as
  `6b340fe0e9` did, turned a picture selection `[p, p+1)` into `[p, p+k+1)` with
  `fGraphics` still set.
- **Refused:** a non-empty bullet or number selection (`fPrefix`, or `sty`
  `styPrefix`), which has no character, so 2.9.244 gives its `cpLim` no defined
  side (an empty one stays before the text like a caret); and a record that
  claims an object and also a frame, a prefix, table cells or a block, whose
  claims need opposite mappings.

Still refused: a CP strictly inside replaced or removed text (body-text
replacement, accepting a deletion, rejecting an insertion), because the
characters it was recorded between no longer exist. Only `add_text` performs pure
insertions, so the body-text facade (non-empty spans only) is unaffected.

### FKP page overflow (`writer/fkp.rs`, `tracked_revision/codec.rs`)

`ChpxFkpBuilder`, `PapxFkpBuilder` and the tracked-revision `build_papx_pages`
charge each property its even-rounded size, but place properties downward from
the count byte at offset 511 and align each start to an even offset; the first
placement can therefore take one byte more than estimated. A page filled exactly
to the estimate put its lowest property one byte into the BX array and
overwrote the last BX (for PAPX pages, the last PHE byte), which the strict FKP
parsers refuse. After the existing estimate chooses a page's entries, the builders
now drop entries while the exact placement would reach the FC/BX arrays. A page
the builders already laid out validly is unchanged, so output changes only where
it was corrupt. A tracked-revision PAPX that only the estimate admitted on an
empty page is now the typed "one PAPX run cannot fit in an FKP" refusal instead
of a corrupt page; the fresh writer's builders keep their "at least one entry per
page" behavior.

### Tests

- `parts::protection::tests`: `word_2002_requires_its_counted_fib_and_dop_shapes`
  (which pinned the removed refusals) becomes
  `word_2002_shapes_are_classified_by_their_protection_fields`: the conforming
  Word 2002 shape, the pre-0759 writer's 500-byte DOP, LibreOffice's FIB and real
  610-byte DOP (`None`), that DOP with an enforced mode 0–3 (`Document`,
  bypassable only with `AllowProtected`), 7 (`None`), 4–6 enforced or not
  (`Unrecognized`), a `DopBase` lock plus key in it (`Document`), an 83-byte DOP,
  and wrong-count, wrong-`cswNew`, truncated-`nFibNew` and truncated-`cswNew`
  FIBs (`Unknown`).
- New: note <11> (`0x00C0`/`0x00C2` read as Word 97, neighbours `Unknown`); Word
  2003 DOPs of 84..599 bytes (`Unrecognized`) and 600..616 and 674 bytes
  (`None`, or `Document` when enforced); every `DopBase` lock, each SHOULD-level
  combination (`Document`), each protection-field MUST (`Unrecognized`), and eight
  non-protection fields set to out-of-domain values (`None`); out-of-bounds
  DOPs at four offsets; and the 35 real fixtures (`None`).
- Kept: 595..615-byte DOPs under 0x0112, `iDocProtCur` 4..6 in a 674-byte DOP,
  zero `lcbDop`, truncated FIB, wrong count, malformed range tables, and the
  `authorize` matrix.
- `doc_tracked_revision_editor`: `picture.doc` edited as checked in; all 35
  readable fixtures accept a tracked insertion at CP 0, reopen and republish
  byte-identically; the three unreadable ones are refused at open; a saved caret
  keeps its CP (`cjklist30`, `empty`, `HeaderFooterUnicode` 0/0/407 → 0/0/412,
  `NoHeadFoot` 179 → 184) and rejecting the insertion restores the recorded
  `Selsf`; the three real FKP-overflow documents reopen after an edit (this test
  fails on `6b340fe0e9`'s code with "malformed CHPX FKP").
- `doc_render_handoff`: `documentProperties.doc` edited as checked in.
- `saved_selection` and `writer::fkp`/`codec` unit tests pin the insertion rule,
  the interior refusals, the exact-capacity CHPX/PAPX pages (overflowing and
  just-fitting) and the typed single-PAPX refusal.

## Per-fixture table (`test-data/ole/doc`)

Producer from the FibRgW97 magic words; "base"/"after" are the
`property_set::Snapshot::protection()` verdicts; the edit columns are
`RevisionEditor::add_text(0, "x")` + `finish` + strict reopen
([probe](results/change-0768/scripts/probe_0768.rs)).

| file | producer | nFib / cbRgFcLcb / cswNew (nFibNew) | lcbDop | base | after | base, insert at CP 0 | after |
| --- | --- | --- | ---: | --- | --- | --- | --- |
| 3endnotes.doc | Word | 0x00C1 / 0x00B7 / 5 (0x0112) | 690 | Unrecognized | None | refused: protection | edits |
| cfb-truncated-final-sector.doc | Word | 0x00C1 / 0x00B7 / 5 (0x0112) | 694 | Unrecognized | None | refused: Selsf CP | edits |
| cfb-v3-uninitialized-size-high-word.doc | other | 0x0068 / 0x1AEE / – | – | Unknown | Unknown | refused: invalid CFB | refused: invalid CFB |
| cjklist30.doc | Word | 0x00C1 / 0x00B7 / 5 (0x0112) | 694 | Unrecognized | None | refused: Selsf CP | edits |
| cjklist31.doc | Word | 0x00C1 / 0x00B7 / 5 (0x0112) | 694 | Unrecognized | None | refused: Selsf CP | edits |
| cjklist34.doc | Word | 0x00C1 / 0x00B7 / 5 (0x0112) | 694 | Unrecognized | None | refused: Selsf CP | edits |
| cjklist35.doc | Word | 0x00C1 / 0x00B7 / 5 (0x0112) | 694 | Unrecognized | None | refused: Selsf CP | edits |
| commented-table.doc | Word | 0x00C1 / 0x00B7 / 5 (0x0112) | 690 | Unrecognized | None | refused: Selsf CP | edits |
| DiffFirstPageHeadFoot.doc | Word | 0x00C1 / 0x00B7 / 5 (0x0112) | 674 | None | None | edits | edits |
| documentProperties.doc | LibreOffice | 0x0101 / 0x0088 / 0 | 610 | Unknown | None | refused: protection | edits |
| duplicate-style-names.doc | LibreOffice | 0x00C2 / 0x006C / 2 (0x00D9) | 600 | Unrecognized | None | refused: protection | edits |
| empty.doc | Word | 0x00C1 / 0x005D / 0 | 500 | Unrecognized | None | refused: Selsf CP | edits |
| endingnote.doc | LibreOffice | 0x0101 / 0x0088 / 0 | 610 | Unknown | None | refused: protection | edits |
| equation.doc | LibreOffice | 0x0101 / 0x0088 / 0 | 610 | Unknown | None | refused: protection | edits |
| FancyFoot.doc | Word | 0x00C1 / 0x00B7 / 5 (0x0112) | 674 | None | None | edits | edits |
| first-header-footer.doc | Word | 0x00C1 / 0x00B7 / 5 (0x0112) | 674 | None | None | edits | edits |
| FloatingPictures.doc | Word | 0x00C1 / 0x00A4 / 2 (0x010C) | 616 | None | None | edits | edits |
| footnote.doc | LibreOffice | 0x0101 / 0x0088 / 2 (0x0101) | 610 | Unrecognized | None | refused: protection | edits |
| HeaderFooterProblematic.doc | Word | 0x00C1 / 0x005D / 0 | 500 | Unrecognized | None | refused: protection | edits |
| HeaderFooterUnicode.doc | Word | 0x00C1 / 0x00B7 / 5 (0x0112) | 674 | None | None | refused: Selsf CP | edits |
| hyperlink.doc | LibreOffice | 0x0101 / 0x0088 / 0 | 610 | Unknown | None | refused: protection | edits |
| image-comment-at-char.doc | Word | 0x00C1 / 0x00B7 / 5 (0x0112) | 690 | Unrecognized | None | refused: Selsf CP | edits |
| inline-endnote-and-footnote.doc | Word | 0x00C1 / 0x00B7 / 5 (0x0112) | 694 | Unrecognized | None | refused: Selsf CP | edits |
| lists-margins.doc | LibreOffice | 0x0101 / 0x0088 / 0 | 610 | Unknown | None | refused: protection | edits |
| Lists.doc | Word | 0x00C1 / 0x00B7 / 5 (0x0112) | 690 | Unrecognized | None | refused: protection | edits |
| NoHeadFoot.doc | Word | 0x00C1 / 0x00B7 / 5 (0x0112) | 674 | None | None | edits | edits |
| PasswordProtected.doc | Word | 0x00C1 / (encrypted) | – | Unknown | Unknown | refused: encrypted | refused: encrypted |
| picture.doc | LibreOffice | 0x0101 / 0x0088 / 0 | 610 | Unknown | None | refused: protection | edits |
| pictures_escher.doc | LibreOffice | 0x0101 / 0x0088 / 0 | 610 | Unknown | None | refused: protection | edits |
| PngPicture.doc | LibreOffice | 0x0101 / 0x0088 / 0 | 610 | Unknown | None | refused: protection | edits |
| table-merged-cells.doc | LibreOffice | 0x0101 / 0x0088 / 0 | 610 | Unknown | None | refused: protection | edits |
| tdf71749_with_footnote.doc | Word | 0x00C1 / 0x005D / 0 | 500 | None | None | edits | edits |
| testPictures.doc | Word | 0x00C1 / 0x00A4 / 2 (0x010C) | 616 | None | None | edits | edits |
| ThreeColFoot.doc | Word | 0x00C1 / 0x00B7 / 5 (0x0112) | 674 | None | None | edits | edits |
| ThreeColHead.doc | Word | 0x00C1 / 0x00B7 / 5 (0x0112) | 674 | None | None | refused: Selsf CP | edits |
| ThreeColHeadFoot.doc | Word | 0x00C1 / 0x00B7 / 5 (0x0112) | 674 | None | None | edits | edits |
| watermark.doc | LibreOffice | 0x0101 / 0x0088 / 0 | 610 | Unknown | None | refused: protection | edits |
| word6-no-table-stream.doc | other | 0x0065 / – / – | – | Unknown | Unknown | refused: pre-Word 97 FIB | refused: pre-Word 97 FIB |

"refused: Selsf CP" files were refused at CP 1 by the protection check, except
`HeaderFooterUnicode.doc` and `ThreeColHead.doc`, which edited at CP 1. Why each
of the 24 was refused (first failing check): the ten LibreOffice files with a
zero `cswNew` failed the FIB check; `footnote.doc` and
`duplicate-style-names.doc` failed the exact DOP length (610 ≠ 594, 600 ≠ 544);
the two Word 97 files failed the typed Dop97 grammar (`Dogrid` display multiples
0); the ten Word 2007+ files failed `Dogrid` or the Dop2010 `docid` 0. The 19
files outside `test-data/ole/doc` are in
[other-doc-table.md](results/change-0768/corpus-probe/other-doc-table.md).

## Adversarial review

- **Can a protected document now classify as unprotected?** Only if a protection
  signal lived in a field the classifier no longer reads. MS-DOC places document
  protection only in the `DopBase` locks and key and in Dop2003 byte 598 (Dop2007,
  Dop2010 and Dop2013 add none), and range protection only in the tables at
  pointers 141–144, which the unchanged `Ranges::parse` reads whenever the
  declared array reaches them. A 0x0101 FIB with 136 pointers has no range tables
  for Word either. Byte 598 is read in every generation whenever present, so a
  LibreOffice-shaped DOP with enforcement on is `Document` (tested), as is a
  crafted 599-byte one.
- **Can the new classifier refuse something the base accepted?** No: for every
  FIB and DOP shape the base accepted (the exact lengths), the new rules read the
  same fields with the same MUSTs, except that `fProtEnabled` with `fLockAtn` or
  `fLockRev` moves from `Unrecognized` to `Document` (both refused by default).
- **Selsf:** the remap result is re-parsed and bounded by the new text length, so
  an invalid record is still `Invalid`; inserting and then rejecting restores the
  recorded bytes (tested on real fixtures). The first version kept an object
  selection's CP (fixed after the review, below).
- **FKP:** the trim only runs when the exact placement overflows, which is exactly
  when the base wrote a page its own parser refuses; every golden and writer test
  is unchanged.

## Independent review and follow-ups

An independent review compared the base and this branch on the 57 DOC files, on
1,730 crafted documents (every `DopBase` lock and combination, the password
hash, `iDocProtCur` 0–7, DOP lengths 83–695, nFib variants, LibreOffice-shaped
FIBs with protection, malformed ranges and encryption), on 148 `Selsf`
insertion scenarios and on 300,000 CHPX and 300,000 PAPX builder
configurations. It found the protection boundary holds: no document goes from
`Document` to `None`, `Unknown` and `Unrecognized` stay refused under
`AllowProtected`, and every `Selsf` rejection restored the recorded bytes. Of
the builder configurations, CHPX pages are byte-identical wherever the base's
were valid and 12,348 the base laid out invalidly are now valid; PAPX pages are
identical in 227,082, fixed in 5,532, and invalid on both sides in 67,386 (a lone
grpprl of 486 bytes or more, which follow-up 2 now refuses). Its verdict was
merge-after-fixes; the fixes are new commits, with no history rewritten.

1. **Object selections (required, `95b485a290`).** The rule is in section "The
   `Selsf` check". Tests: each flag (`fGraphics`, `fShape`, `fFrame`, `fPrefix`,
   `styPrefix`), empty object and prefix records, and the four conflicting
   combinations; and a fixture test that points the `Selsf` at real inline
   pictures (testPictures CP 8, FloatingPictures CP 4582) and shape anchors
   (testPictures CP 100, FloatingPictures CP 5793, image-comment-at-char CP 3),
   inserts there, checks that the selection still covers exactly that
   character, and rejects the insertion back to the recorded bytes.
2. **FKP builders on caller input (`0792812830`).** The public `PapxFkpBuilder`
   placed a lone PAPX even after its size estimate failed: a grpprl of 486–507
   bytes overlapped the BX and FC arrays, and from 508 bytes `511 − (len +
   extra)` underflowed. The fresh writer reaches it: a table row keeps its table
   properties in its row-end PAPX (22 bytes a column), and at the base
   `Writer::add_table(1, n)` wrote rows of up to 21 columns correctly, wrote a
   file the reader rejects at 22 ("malformed SPRM sequence"), and panicked in
   `write_to` from 23 to 63 columns (`fkp.rs:353`, subtract with overflow). Both
   builders now refuse, with `InvalidInput`, a property no page can hold, and
   place entries with checked arithmetic; the CHPX builder also refuses a
   grpprl over 255 bytes, which `Chpx.cb` cannot count, instead of cutting it to
   255 bytes. A scan of the other page builders found no arithmetic that caller
   input can underflow: the tracked-revision PAPX builder bounds runs at 510
   bytes, places with `checked_sub` and refuses a lone oversized run; CHPX runs
   are checked before the builder; `pack_dttm` bounds its year; the piece-table
   splices take validated ranges. Tests: PAPX 485 fits (BX word offset 11); 486,
   487, 507, 508, 509, 510, 511 and 4,096 are refused, as is an oversized entry
   after a small one; CHPX 255 fits and 256 is refused; the fresh writer writes
   21 columns and refuses 22, 23, 40 and 63. Writer goldens are unchanged.
3. **One classifier (`83fa8563d7`).** The positional body-text source
   (`body_text::source`) kept its own copy of the pre-0768 grammar (exact
   `cswNew` and `cbRgFcLcb` per generation, the typed DOP parser, exact DOP
   lengths), so a LibreOffice FIB failed at open and unusual DOPs were
   `Unrecognized`. The shared classifier gains `classify_with`, which takes the
   DOP and, only when a range-protection pointer declares data, the table
   stream from the caller; `classify` wraps it. The source reads the complete
   FIB and feeds its bounded readers to it, keeping its non-protection refusals
   (encryption, macros, fields, drawings, revision authors) and its refusal of
   revision marking in an unprotected document. A FIB or DOP shape the
   classifier cannot prove, including a missing DOP, now opens read-only as
   `Unknown` or `Unrecognized` (changed publication refused under every policy)
   instead of failing to open, and nFib 0x00C0 is read as Word 97, as the
   classifier reads it. Agreement tests: on all 57 DOC files, 49 identical
   verdicts (`None`), 7 encrypted or pre-Word 97 files refused by the source
   before it reads protection (shared verdict `Unknown`), and 1 invalid CFB
   unreadable by both; on the review's 1,730 crafted records, 1,708 identical
   verdicts (591 `Document`, 510 `None`, 507 `Unrecognized`, 100 `Unknown`), 20
   encrypted-flag variants refused by the source's encryption gate, and 2
   wrong-`cswNew` FIBs that break its `fcMin` check where the shared verdict is
   `Unknown`. The source's tests use `documentProperties.doc` as checked in; the
   helpers that rewrote its FIB and DOP for the old grammar are gone.
4. **Insertion after the final paragraph mark (`de2df97e1f`, pre-existing).**
   `add_text` accepted `cp == ccpText`, writing text after the main story's
   final paragraph mark, which MS-DOC 2.3.1 requires to be the story's last
   character. No editor in the crate defines append-after-the-mark semantics,
   and moving the caller's CP would be a guessed edit, so it is refused like a
   CP beyond the story; the documentation points to `ccpText − 1`. Every
   existing caller inserts at CP 0. Test on NoHeadFoot.doc (ccpText 180): 180
   and 181 are refused with no change, and 179 extends the last paragraph,
   which still ends with its mark after reopening.
5. **`fStyleLock` and `fStyleLockEnforced`** stay outside the classifier, as
   recorded under "Findings outside this change".

## Measurement (control; no effect expected)

The classifier runs when the DOC editor opens, inside the timed region of the
0729/0730 public lifecycle probe (open, edit, replace paragraph 0, commit, copy
output) and outside the harness's `doc_semantic_one_edit_save` interval
(`Snapshot::from_bytes` precedes it). The FKP check runs whenever CHPX pages are
rebuilt: every body-text replacement, tracked insertion and fresh DOC write.
Both legs were built with identical commands (`cargo build --release --locked
--offline`, Rust 1.95.0) from the base and the branch, staged at equal-length
paths, and run with identical arguments, pinned with `taskset -c 6` on an AMD
EPYC 9R45 (32 cores) shared with other agents (1-minute load average 23–25
during batch 1 and 8–9 during batch 2). Every probe sample's output equals the
expected output, and both legs produce identical bytes: NoHeadFoot
`5d79970a…`, FloatingPictures `6ba1bc97…`, harness corpora `c9e22d55…` (tiny)
and `346cb6e8…` (large).

### Exact instructions of the timed region (callgrind, both legs, identical arguments)

| case | before | after | change |
| --- | ---: | ---: | ---: |
| probe `docnohf` (NoHeadFoot.doc), Ir per lifecycle | 1,624,955 | 1,618,967 | −0.37% |
| probe `docfloat` (FloatingPictures.doc), Ir per lifecycle | 22,093,511 | 21,909,149 | −0.83% |
| harness `doc_semantic_one_edit_save` tiny, timed calls over one process | 6,381,797 | 6,346,911 | −0.55% |
| harness `doc_semantic_one_edit_save` large, timed calls over one process | 44,936,553 | 45,056,754 | +0.27% |

`classify` (inclusive, whole process): tiny 144,694 → 27,966 Ir, large 42,223 →
8,058. `ChpxFkpBuilder::generate_pages`: tiny 7,078 → 7,541, large 82,814 →
97,688 (the exact-placement check). Of the large case's +120K Ir (+0.27%),
15K are that check and the rest lie in unchanged code, consistent with inlining
differences (`Editor::put_streams_shared_with_rendered` is out of line only
after); see
[counters](results/change-0768/counters/harness-large-inclusive-diff.txt).

### Wall clock, ABBA

Two independent batches of eight rounds (16 processes per case per batch).
Probe processes: 5 warmups + 40 samples (`docnohf`) or 30 (`docfloat`); harness
processes: 5 warmups + 30 samples, both writer shapes in one process. Ratios are
after/before of the paired per-round process p50s (geometric mean, 95% bootstrap
CI over rounds).

| case | batch | before, median of p50s | after | paired p50 ratio [95% CI] |
| --- | --- | ---: | ---: | --- |
| probe `docnohf` | 1 | 112,588 ns | 109,458 ns | 0.990 [0.962, 1.015] |
| probe `docnohf` | 2 | 108,551 ns | 108,348 ns | 0.992 [0.966, 1.017] |
| probe `docfloat` | 1 | 1,085,071 ns | 1,185,129 ns | **1.072** [1.013, 1.128] |
| probe `docfloat` | 2 | 1,050,078 ns | 1,013,562 ns | 0.994 [0.940, 1.062] |
| harness tiny | 1 | 40,818 ns | 39,555 ns | 0.967 [0.953, 0.980] |
| harness tiny | 2 | 40,580 ns | 39,113 ns | 0.962 [0.950, 0.970] |
| harness large | 1 | 851,870 ns | 660,606 ns | **0.785** [0.768, 0.802] |
| harness large | 2 | 824,959 ns | 655,636 ns | **0.804** [0.792, 0.820] |

### The harness `large` shape runs in two heap states

The same binaries run the `large` shape at about 650 µs with ~55.8K page faults
per process or at 810–850 µs with ~74K. Which state a process lands in depends
on its allocation history and environment, not on this change's work: with both
shapes in one process the after leg is 20% faster; alone, it was 27–31% slower
in the first sequential observation (before 652/647 µs, 55.8K faults; after
831/846 µs, 73.5K faults) and 2.9% slower in a later ABBA (both legs in the slow
state, 74.5K vs 75.6K faults). With glibc's dynamic thresholds pinned
(`GLIBC_TUNABLES=glibc.malloc.mmap_threshold=8388608:glibc.malloc.trim_threshold=67108864`)
both legs run at ~620 µs with equal page faults (32.8K vs 32.9K), paired
ratios 0.990–1.010, median 1.000. The call sites and call counts of every Vec
growth are identical in both legs; glibc's realloc does more work in the after
leg's large run because the heap layout differs
([callers](results/change-0768/counters/harness-large-vec-growth-callers.txt)).
One plausible trigger is that the new `classify` no longer allocates the DOP
copies the typed parser made; this is not proven. Record 0757 saw the same kind
of glibc trim/mmap effect on the XLS writer.

### Regression flags (>5%)

- probe `docfloat`, batch 1: +7.2% p50 (CI 1.013–1.128) under a load average of
  23–25; batch 2 (load 8–9): −0.6%, CI 0.940–1.062. Exact instructions: −0.83%.
  Not reproduced; recorded.
- harness `large`, `--writer-shape large` alone: +27–31% (two processes per leg,
  sequential) and +2.9% (four ABBA pairs), with page faults tracking the time;
  −20% in both ABBA batches with `tiny,large`. Exact instructions: +0.27%; with
  malloc thresholds pinned: 1.000. Attributed to glibc heap state, recorded.
- Process-level `perf` counts are not a timed-region measure here: they include
  kernel page-fault handling, and ~96% of a harness process is SHA-256 over its
  own executable (`current_executable_identity`).

## What is not claimed

- No performance claim; the controls show that the classifier's cost change is
  not measurable in wall clock against glibc heap-state effects.
- No claim that any real protected DOC was tested: no file in the repository's
  57 carries a protection signal. Protected states are covered by synthesized
  DOPs built from real fixtures' DOP bytes.
- No change to what a DOC reader exposes: the typed DOP API still validates every
  field; only the protection classifier stopped using it.
- No native Word or LibreOffice round trip of the edited files.
- The `Selsf` rule is spec-conformant, not Word-observed: Word itself rewrites
  `Selsf` on every save.
- The review follow-ups were not re-measured. On the measured paths they add a
  size scan and one page-fit check per CHPX FKP page (no allocation) and a flag
  test in the `Selsf` remap; the positional body-text source, which now reads
  the complete FIB and no longer parses the typed DOP at open, is not on a
  measured path.

## Findings outside this change (not fixed)

- `validate_fib_shape` checks `cswNew × 2` bytes starting at `cswNew` itself
  rather than after it, so a `FibRgCswNew` missing its last word passes. It
  carries no protection data; tightening it would make the synthetic FIBs of
  seven package test modules (13 tests, which write only `nFibNew`) `Unknown`.
  Kept as at the base.
- Dop2003 `fStyleLockEnforced` (Word's "limit formatting to a selection of
  styles") is not treated as an editing restriction by the base or this change;
  the direct-formatting edits (bold) could conflict with it.
- `test-data/office-interop/litchi-changed/noheadfoot-litchi.doc` has a `Selsf`
  beyond its text (cpFirst 179 > ccpText 135): an earlier litchi edit did not
  remap it. It is refused as invalid, correctly.
- Other integration tests still normalize real fixtures with
  `with_valid_word97_dop` (`doc_body_transaction`, `user_defined_hyperlinks`,
  `binary_digital_signatures`, `doc_captions`, `dofr_edits`,
  `doc_mtef_equation_writer`, `saved_selection`, `vba_signatures`); several could
  now use the fixtures as checked in.
- The fresh writer refuses, with a typed error, a table row whose table
  properties exceed one PAPX FKP page (a plain row of 22 or more cells).
  Writing it needs `sprmPHugePapx`, which stores the properties in the Data
  stream as `PrcData`; litchi's reader expands it, but the tracked-revision
  editor's PAPX rewrite reads grpprls from the FKP pages and does not model it.
- An object selection whose object is removed (accepting a deletion, or
  rejecting an insertion, of exactly that character) collapses to an empty
  range at the removal point with its object flag still set, as at the base. It
  covers no text, so it claims nothing false about content; refusing it would
  newly refuse those edits.

## Verification

All gates pass on `83fa8563d7`, the last code commit, except the known
`non_iwork_gate.py verify` failure (`litchi-xldm` inventory), which fails
identically at the base; details in [gates.txt](results/change-0768/gates.txt):
`cargo fmt --all --check`; `cargo check` of `litchi-doc` (all targets) and the
facade with every format feature; warning-denied Clippy of `litchi-doc` (lib
and all targets) and rustdoc; `cargo test -p litchi-doc` (1326 passed, 13
ignored) and the facade with `doc,docx,ppt,pptx,xls,xlsx,xlsb,odt` (382
passed, 7 ignored); crate boundaries; structural perf-claim check. Every commit
was built and tested on its own: 1297, 1301 and 1306 passed for the original
change, then 1311, 1312, 1315 and 1326 for the review follow-ups. The harness
was not changed.

## Cleanup

Binary identities are in [binaries.sha256](results/change-0768/binaries.sha256).
Removed: `targets/0768` (≈45 GB: debug test, Clippy and rustdoc builds and the
release harness and probe), `targets/0768-before` (≈4 GB), the detached base
worktree `0768-before-src`, and the scratch directory (staged binaries, probe
copies, callgrind profiles, raw runs; the raw JSON reports are kept here
gzipped). The debug trees were deleted mid-task when the shared disk reached 0
bytes free. After the review follow-ups, the rebuilt `targets/0768` (gate and
test builds) and the scratch directory were removed again once this record was
committed. The worktree and branch are kept. See
[cleanup.json](results/change-0768/cleanup.json).
