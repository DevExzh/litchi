# 0640: two DOC field-table refusals were reading a redundant annotation as a structural fact; the five duplicate-style refusals are correct and already have a documented lenient path

Status: retained, correctness fix. `performance_claim: none`. Two refusals in
`litchi-doc`'s `Plcfld` reader are removed because the bits they check restate
what the `FieldList` grammar already fixes and real producers leave them stale;
five stylesheet refusals are confirmed correct and unchanged. No timing was
taken and none is claimed.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## Why this record exists

Change 0587's survey listed, as correctness defect 4, seven `.doc` fixtures the
eager facade route refuses: two for a field-table consistency disagreement
(`ole/doc/watermark.doc` on `grffldEnd.fNested`, `poi/test-data/document/test.doc`
on `grffldEnd.fHasSep`) and five for `style names and aliases must be unique`.
0587 called the first two "ordinary Word files" and left the second five
"unverified whether [leniency] admits these". This record settles all seven
against the fixture bytes.

The route in question is the one change [0609](0609-facade-doc-source-route-design.md)
measured and named Route E: `litchi::Document::open` slurps the file and calls
`doc::Package::from_ole_file` then `package.document()`
(`crates/litchi/src/document/doc.rs:928-947`). The differential below drives
exactly that pair, so its admission set is the facade's.

## What was changed

`crates/litchi-doc/src/parts/fields/package.rs`, `build_fields`: the two
consistency refusals

```rust
if flags.has_separator != has_separator {
    return Err(corrupted("grffldEnd.fHasSep disagrees with the FieldList"));
}
if flags.nested == stack.is_empty() {
    return Err(corrupted("grffldEnd.fNested disagrees with field containment"));
}
```

are removed. Nothing else in the reader moves. Every other `Plcfld` refusal
stays: the allocation and marker-count bounds, the byte-length shape, CPs
strictly increasing, marker CPs inside the story's character count, the
field-character tag whitelist, a separator outside a field, two separators in
one field, an unmatched end marker and an unmatched begin marker.

Three documentation changes accompany it: a comment at the removal site giving
the reasoning and naming both witness fixtures; doc comments on
`FieldEndFlags::nested` / `::has_separator` saying they are the file's
annotation as written and may disagree; and doc comments on `Field::end_flags`
(as written), `Field::nesting_depth` and `Field::has_separator` (derived from
the grammar).

Two tests change in `crates/litchi-doc/src/parts/fields/tests/parse.rs`. The
malformed matrix loses the case that was only invalid because of these bits
(`plcf(&[1, 3, 5], &[[0x13, 0x21], [0x15, 0x80]])`, a separator-less field
whose end sets `fHasSep`) and gains two that are structurally invalid instead
(an end marker that closes nothing inside an otherwise balanced sequence; a
separator belonging to no open field). A new test,
`stale_end_flags_are_read_from_the_structure_and_preserved`, builds both witness
shapes and asserts that they parse, that `nesting_depth` and `has_separator`
come from the grammar, that `end_flags` keeps the stale bits, that
`to_plcf_bytes` replays the input byte for byte, and that an unmatched begin is
still refused.

## What was not changed

The five duplicate-style refusals. `validate_styles`
(`crates/litchi-doc/src/parts/styles/semantic/validation.rs:195-207`) is
correct as written and stays.

## Why it is sound

### The two bits are redundant, and the structure is the normative carrier

`Plcfld` (MS-DOC 2.8.25) stores a CP array and a parallel array of two-byte
`Fld` descriptors (MS-DOC 2.9.88-2.9.90, as `crates/litchi-doc/src/parts/fields/mod.rs`
cites them). The begin / optional separator / end sequence *is* the field
structure: the reader derives `separator_cp` from the separator marker it
actually saw inside the field, and `nesting_depth` from how many fields are
still open. `grffldEnd.fHasSep` and `grffldEnd.fNested` restate those two facts
and nothing else. Reading them cannot resolve an ambiguity, because there is no
ambiguity to resolve; it can only contradict a structure that was already
validated.

A disagreement is therefore not evidence about the *fields*; it is evidence
about the producer. Everything that would indicate a damaged table — CPs out of
order, a CP past the story's character count, an unknown field-character tag, an
unbalanced marker sequence, two separators in one field — is checked separately
and is untouched here.

**What the removal does give up.** A disagreement can still indicate one narrow
class of damage the grammar does not see: a separator marker dropped from an
otherwise balanced field (`fHasSep` would stay set), or an enclosing begin/end
pair lost together (`fNested` would stay set on the field that was inside it).
The check cannot distinguish that from a stale annotation — POI's `test.doc` is
byte-for-byte the shape a dropped separator would produce — and this corpus
supplies the false positive twice while supplying no example of the true
positive. What is lost if such a file ever appears is bounded and visible: the
field reads as separator-less, so its result text is reported inside
`code_range()` rather than as `result_range()`, and `end_flags.has_separator`
still carries the producer's contradicting claim for a caller that wants to act
on it. What was lost before was the entire document, every time, including for
the two files in this corpus.

### Error identity and preservation

Both bits are still parsed, still stored on `FieldEndFlags`, and still written
back by `FieldDescriptor::to_bytes`, so `FieldStoryTable::to_plcf_bytes` remains
byte-identical for any table it read; the new test asserts this on both witness
shapes. A caller that cares what the producer claimed reads `end_flags`; a
caller that wants the truth reads `has_separator` and `nesting_depth`. Nothing
is repaired, normalized or silently corrected, so this is preservation, not
leniency, and no `ToleranceReport` entry is warranted.

The write side is unaffected. `crates/litchi-doc/src/writer/fields.rs:205`
builds a fresh `Plcfld` for an authored document and derives both bits from the
structure it is emitting:

```rust
let flags = u8::from(has_separator) << 7 | u8::from(!field_stack.is_empty()) << 6;
```

That is ADR 0006's division exactly — readers preserve real-world quirks,
writers do not bake quirks into semantic models — and it is why this change is
read-side only.

### What the bytes say

Both fixtures were decoded straight from the Table stream by the retained probe
(`probe/main.rs`, mode `fld-dump`), which walks the `Plcfld` independently of
`litchi-doc`'s field parser and annotates each end marker.

`ole/doc/watermark.doc` (nFib 0x0101, `1Table`, `fcPlcfFldMom` = 4554,
lcb = 112, 18 markers) is a single `TOC` field (`flt` 0x0D) containing five
`HYPERLINK` fields (`flt` 0x58), each with its own separator. Every one of the
five inner end markers carries `grffldEnd = 0x80`: `fHasSep` set and correct,
`fNested` clear although the marker closes a field nested one level inside the
`TOC`. The outer `TOC` end marker also carries 0x80 and is correct, because it
is not nested. The marker sequence is balanced and complete
(`residual_open = 0`), the CPs are strictly increasing, and the field
instruction text reads normally:

```
field #2..#4   cp 73..128  depth=1  " HYPERLINK  \l \"__RefHeading__8_476954814\"\x01\x14\x0b1. TERM\t1"
field #0..#17  cp 66..413  depth=0  " TOC \x14 HYPERLINK ..."
```

`poi/test-data/document/test.doc` (nFib 0x00C1, `1Table`, `fcPlcfFldMom` = 538,
lcb = 16, 2 markers) has exactly one field, ` SEQ CHAPTER \h \r 1`, with a begin
at CP 0 and an end at CP 21 and no separator marker and no separator character
in the text between them. Its end marker carries `grffldEnd = 0x80`, so
`fHasSep` is set on a field that has no separator. A `SEQ` field with the `\h`
switch has no result by definition, so a missing separator is exactly what the
instruction text calls for; the field instruction is syntactically complete and
the PLCF's framing is exact.

### How common the disagreement is

`structure/fld-corpus-scan.txt` is the same independent walk over all 57 `.doc`
fixtures under `test-data/`: 98 field end markers in total.

| bit | agrees | disagrees |
| --- | ---: | ---: |
| `fNested` clear, no enclosing field | 93 | — |
| `fNested` set, enclosing field present | 0 | — |
| `fNested` clear, enclosing field present | — | 5 |
| `fHasSep` matching the separator's presence | 97 | — |
| `fHasSep` set, no separator | — | 1 |

`fNested` is set on **none** of the 98 end markers in this corpus, including the
only five that are genuinely nested. A check whose positive evidence is 0 out of
5 on the only nested fields available is not distinguishing damaged files from
sound ones. `fHasSep` does track the structure in 97 of 98 cases, which is why
that half of the change rests on the redundancy argument and on the SEQ field's
own text rather than on a population argument.

### Correcting 0587 on provenance

`structure/provenance.txt` is the retained evidence for this section.

0587 called both files "ordinary Word files". `watermark.doc` is not: its
stylesheet carries `WW8Num2z0`-`WW8Num3z8` character styles, `Internet Link`,
and `Absatz-Standardschriftart` beside a separate `Default Paragraph Font`,
and its bookmarks are named `__RefHeading__N_...` — the OpenOffice.org /
LibreOffice `.doc` export signature. Its `\x01CompObj` says
`Microsoft Word-Dokument` / `Word.Document.8`, which an exporter writes too.
`poi/.../test.doc` has no `SummaryInformation` stream at all and its stylesheet
(`Default Paragraph Font`, `Default Para`, `Default Par1`..`Par6`) is consistent
with a Word 97 file, but nothing in the file identifies the producer. **This
record does not claim that Microsoft Word writes either file**, and the
disposition does not depend on it: the argument is that the bits are redundant
with a structure that is intact in both files, that an independent consumer
reads both (LibreOffice 26.2.5.2 converts both to text without complaint), and
that ADR 0006 requires a reader to preserve real-world quirks rather than refuse
over them. Under the brief's reading of ADR 0006 — refusals must be correct, not
merely conservative — these two were merely conservative.

### The five duplicate-style refusals are correct

All five fixtures (`duplicate-style-names.doc`, `footnote.doc`,
`lists-margins.doc`, `picture.doc`, `pictures_escher.doc`) carry the identical
defect: `istd` 10, the fixed-index `sti` 65 built-in, and `istd` 15, a user
character style, both named `Absatz-Standardschriftart`. `structure/stsh-*.txt`
has the full sheets. MS-DOC 2.9 requires a style's name and each alias to be
unique within the stylesheet, as `validate_styles` and
`crates/litchi-doc/src/leniency.rs` both already state, so these files violate a
stated requirement and `Leniency::Strict` is right to say so. They are
OpenOffice.org/LibreOffice exports: the same sheets carry `WW-`-prefixed style
names, and `footnote.doc` mixes Russian built-in names with a German
`Absatz-Standardschriftart`.

**The question 0587 left open has a clean answer: yes.**
`OpenOptions::default().with_leniency(Leniency::TolerateStylesheetDefects)`
already admits all five, on the unmodified base commit.
`structure/open-leniency.txt` is the run:

| fixture | strict | lenient |
| --- | --- | --- |
| `duplicate-style-names.doc` | refused | OK, 290 bytes of text, 16 paragraphs |
| `footnote.doc` | refused | OK, 59 bytes, 7 paragraphs |
| `lists-margins.doc` | refused | OK, 175 bytes, 4 paragraphs |
| `picture.doc` | refused | OK, 1,772 bytes, 1 paragraph |
| `pictures_escher.doc` | refused | OK, 1,014 bytes, 12 paragraphs |
| `watermark.doc` | refused | refused, identically |
| `poi/.../test.doc` | refused | refused, identically |

Each lenient read reports exactly one tolerated defect,
`DuplicateStyleName` at style index 15, and nothing else. Leniency correctly
does not touch the two field-table refusals, because `Leniency` is scoped to
stylesheet defects by construction.

## Measured

No timing. The brief asked only for counts confirming that the open does not
move on the fixtures that already open; `counts/counts-summary.txt` has them.

Syscall census for one eager open (`strace -c`), before and after, on the
86.5 KB, 67-field UTAS travel form and on `FloatingPictures.doc`: **identical** —
5 `openat`, 14 `read`, 14 `lseek`, 2 `pread64`, 5 `close`, 16 `mmap` on the
first; 5 `openat`, 19 `read`, 19 `lseek`, 2 `pread64` on the second. The change
touches no I/O, and the census confirms it.

Instructions, callgrind isolation pair (10 and 110 opens differenced, divided by
100), `RAYON_NUM_THREADS=1`, CPU 16:

| fixture | end markers | before Ir/open | after Ir/open | delta | delta % |
| --- | ---: | ---: | ---: | ---: | ---: |
| UTAS travel form | 67 | 4,153,320 | 4,150,518 | -2,802 | -0.0675% |
| `FloatingPictures.doc` | 9 | 2,967,375 | 2,955,382 | -11,992 | -0.4041% |

**These deltas are layout drift, not the removed work, and nothing is claimed
from them.** The change removes two boolean comparisons per end marker, at most
about 300 instructions on the 67-marker fixture; both deltas exceed that. A
first pass of the same pair through a driver binary that differed only by one
unused extra subcommand put the UTAS delta at **+339 Ir (+0.0082%)** rather than
-2,802 — the same `litchi-doc` change, same fixtures, opposite sign. The removal
is below this measurement's layout floor.

## Correctness evidence

**Corpus differential, all 57 `.doc` fixtures, change 0596's oracle verbatim.**
`differential/digest-{before,after}.tsv` carry one line per fixture: open
outcome or the exact `Debug` text of the refusal, `Document::text()` length and
hash, paragraph count, `paragraph_count()` from the separate public query, a
hash over every paragraph's `Debug` form (which carries its resolved properties,
its runs and each run's character properties and revision marks), the section
count and a hash of the section table, a hash of `get_all_subdoc_ranges()`, and
the length and hash of `FileInformationBlock::raw_data()`.

The before digest is **byte-identical to change 0596's retained after digest**
after path normalization, which cross-checks both the oracle and the claim that
nothing between 0596 and this base moved any DOC observation.

`differential/digest-diff.txt` is the whole before/after difference after the
absolute `test-data` prefix is normalized: **two changed records, both on
fixtures that were already refused and still are.** 42 fixtures are admitted on
each leg and 15 refused on each leg; every one of the 42 is identical in every
column, and no fixture changed admission status.

**The two witness fixtures still do not open — for different reasons.** Removing
the field-table refusals unmasks a further refusal on each:

| fixture | before | after |
| --- | --- | --- |
| `ole/doc/watermark.doc` | `grffldEnd.fNested disagrees with field containment` | `bookmark ibkl values must be unique and in range` |
| `poi/.../test.doc` | `grffldEnd.fHasSep disagrees with the FieldList` | `malformed SPRM sequence: truncated SPRM opcode at byte 3: 1 byte(s) remain` |

Those two refusals are outside this brief's seven and are **not** changed here;
see Limitations for what was established about them and why they are left.
Because nothing is newly admitted, the plausibility check the brief prescribes
has no subject on the field-table side.

**Plausibility of the lenient stylesheet path.** The check was run instead on
the five fixtures the existing leniency admits, to substantiate the answer that
their refusal is correct *and* that the documented escape hatch yields a real
document rather than a damaged one. `text/plausibility.txt` compares
`Document::text()` under `TolerateStylesheetDefects` against LibreOffice
26.2.5.2's `txt:Text (encoded):UTF8` export of the same file, on normalized word
sequences. `duplicate-style-names.doc` and `picture.doc` are identical. The
three differences are all known category differences between a stored-text
reader and a rendering exporter and nothing else: LibreOffice adds
`lists-margins.doc`'s four generated list labels (every one of litchi's 32 words
appears in LibreOffice's text); litchi returns `pictures_escher.doc`'s stored
`HYPERLINK` field instructions where LibreOffice returns only the results
(105 of 109 words match, the four others being the field keyword and two
instruction-plus-result URL concatenations); and litchi returns
`footnote.doc`'s footnote, endnote and comment stories, which a body-text export
does not carry.

**Gates** (`gates.txt` has each tail), run in the candidate worktree:

| gate | result |
| --- | --- |
| `cargo fmt --all --check` | pass |
| `cargo clippy -p litchi-doc --all-targets` | pass, no warnings |
| `cargo test -p litchi-doc` | pass; 41 test binaries, 1,178 passed, 0 failed, 13 ignored |
| `cargo doc -p litchi-doc --no-deps` | pass |
| `cargo clippy -p litchi --features docx,xlsx,pptx,xls --all-targets` | pass; 6 warnings, all pre-existing |
| `cargo test -p litchi --features docx,xlsx,pptx,xls` | pass; 26 test binaries, 266 passed, 0 failed, 7 ignored |
| `cargo test --locked` in `tools/perf-baseline` | pass; 19 test binaries, 531 passed, 0 failed, 1 ignored |

The feature-bearing `litchi` suite and the harness's own suite are the two gaps
the first wave found; both were run because `litchi-doc` is a shared crate. The
six `litchi` clippy warnings under that feature set (`missing_ooxml_catalog_part_error`
never used, one needless `mut`, four `passing a unit value to a function`) are
**pre-existing**: the same clippy invocation on the untouched
`before-c7326f680` checkout emits the identical six, and none is in
`litchi-doc`. They are warnings rather than denied lints on this feature
combination, so the gate exits 0 on both legs.

## Validation preserved

Every structural `Plcfld` refusal listed under "What was changed" still fires,
and the new matrix cases prove two of them directly. No limit was weakened:
`MAX_PLCFLD_BYTES` (64 MiB) and `MAX_FIELD_MARKERS` (1,000,000) are untouched,
as are the byte-length shape check, the CP-array overflow check and the
serialization-size check. No `unsafe` was added. Nothing outside
`crates/litchi-doc/src/parts/fields/` changed. Validation still does not mutate:
the removal deletes two comparisons and stores no substitute. `Leniency` gains
no variant and `StylesheetDefect` gains no member, so the closed lenient-repair
contract is exactly as it was. Byte-stable writing is unaffected in both
directions — a table read from a file replays its original bytes, and an
authored table still derives both bits from the structure it emits.

## Limitations — what is not claimed

- **No fixture is newly admitted by this change.** A file whose only defect is a
  stale `grffldEnd` bit would now open, and this corpus contains no such file,
  so the reach gain is **asserted, not witnessed**. A removed refusal is not an
  admission until a file opens, and none did.
- **No producer claim.** This record does not assert that Microsoft Word writes
  either disagreement; `watermark.doc` is an OpenOffice.org/LibreOffice export
  and `poi/.../test.doc`'s producer is unidentified. 0587's "ordinary Word
  files" is corrected above.
- **No spec quotation.** No copy of [MS-DOC] is retained in this repository, so
  the specification is cited only as the crate's own comments cite it
  (2.8.25 for `PlcFld`, 2.9.88-2.9.90 for `Fld`/`grffldEnd`, 2.9 for the
  stylesheet name-uniqueness requirement). Whether [MS-DOC] states a MUST for
  `fNested`/`fHasSep` consistency is **unverified**; the disposition rests on the
  bits being redundant with a validated structure, not on the absence of a MUST.
- **No performance claim.** The instruction deltas above are reported as
  evidence that the open did not move materially, and are explicitly attributed
  to code layout. No timing leg was run and no A/A floor was measured, because
  nothing is claimed that would need one.
- **The two unmasked refusals are not dispositioned.** A bounded investigation
  (retained in `follow-ups.md`) decoded both and reached a *provisional* reading
  that each is also stricter than the format — `watermark.doc`'s repeated
  `FBKF.ibkl` = 4 resolves to `PlcfBkl[4] == PlcfBkl[5] == 535`, so the unused
  and duplicated slots hold the same CP; and `test.doc`'s failing grpprl is a
  4-byte `PapxInFkp` run whose `cb == 0` even-length encoding overshoots a
  3-byte SPRM by one `0x00` pad byte. Neither verdict is acted on here. Both
  touch different subsystems (`parts/bookmarks.rs`, `sprm.rs` /
  `parts/pap_bin_table.rs`), both would move when a refusal happens, and both
  rest on readings this record could not verify against the specification. They
  are queued, not decided.
- **Five fixtures, one defect.** The duplicate-style disposition is established
  on five files that share a single producer and a single duplicated name. A
  different duplicate-name shape is not covered by this evidence.

## Retained evidence

[`results/change-0640/README.md`](results/change-0640/README.md).
