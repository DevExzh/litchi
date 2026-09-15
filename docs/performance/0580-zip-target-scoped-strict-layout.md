# 0580: a ZIP strict-layout proof scoped to the member being read

Status: retained. `performance_claim: none` — this record carries deterministic
read and byte counts and a **semantic change**. No timing is measured and none
is claimed.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements candidate **(b)** of change
[0575](0575-zip-lazy-strict-layout-design.md), which froze the design, the
adversarial witness and a numeric prediction before anything was measured. The
project owner approved (b) as the default on the record. Change 0575's
disposition section records the outcome; the scoring is below, and **three of
its twelve read predictions are falsified**, one of them by a soundness defect
in 0575's own residual-window derivation.

## What was removed

`build_strict_layout_proof` in `soapberry-zip` — both copies, the slice-backed
`ArchiveReader` one and the source-backed `IndexedArchive` one — validated
**every** central record of the archive, in local-header order, before returning
the layout of the one member the caller asked for:

```rust
for layout_entry in &self.layout {
    let span = self.archive.validate_strict_entry_layout(/* … */)?;   // one read each
    if previous_end.is_some_and(|end| span.local_header_offset < end) { /* refuse */ }
    // …
}
```

On `ConditionalFormattingSamples.xlsx` that is 132 positional reads and 10,740
bytes of local-header framing to deliver one 1,251-byte member.

The loop is replaced by a proof scoped to the target:

1. **Distinct local-header offsets**, over every record, at **zero I/O**. A
   record's local-header offset is copied verbatim from its central record, so
   this is a pure central-directory property. It keeps its own refusal.
2. **The target's own layout**, in full and unchanged, through the same
   `validate_reader_entry_layout_with_name_policy` change 0573 left, with the
   same `next_local_header_offset` read hint.
3. **Successors**, at **zero I/O**. The layout is sorted by local-header offset
   and the offsets are now known distinct, so the record that follows the target
   has the smallest offset of any later record, and a later record cannot reach
   backwards past its own offset. One `u64` comparison settles every successor.
4. **Predecessors**, pruned by a zero-I/O central-directory bracket, and
   otherwise settled by **one 30-byte read** of that record's fixed local
   header — no variable region, no name, no CRC, no method, no size
   cross-check. A record whose bracket already places its payload past the
   target's offset is refused with no read at all.

Two error identities were removed because they became unreachable:

| Removed | Why |
| --- | --- |
| `"strict streaming local span extends into the central directory"` | The post-loop archive-wide restatement of a per-entry check. Change 0575 §"What the proof establishes today" showed it was already dead code: `validate_reader_entry_layout_with_name_policy` refuses `local_end`, `data_end_offset` and `span_end` past the directory offset per entry, so every span in the proof satisfied the post-loop predicate before it ran. That per-entry check is untouched and still runs for the target, and a predecessor whose span runs past the directory offset necessarily covers the target and is refused as an overlap. |
| `"strict layout proof publication failed"` | A defensive branch that re-read the cache state immediately after setting it to `Ready`. The new publication path returns the proven layout directly. |

The retained `Vec<StrictEntryLayout>` plus `HashMap<u64, usize>` for all *n*
records is replaced by a per-record memo holding only touched records, plus one
`u64` per record of prefix-maximum span bracket used to stop the predecessor
scan. Four allocation labels change accordingly: `"archive reader strict layout
proof"`/`"map"` and `"indexed archive strict layout proof"`/`"map"` become
`"strict layout target memo"`, `"strict layout span memo"`, `"strict layout
span bracket"` and `"strict layout neighbour spans"`. Every one of them is still
a `try_reserve` on a named resource.

## The semantic difference, in bytes

**Before**, reading any member required that *every* central record in the
archive validate its own local layout and that *no two* records' declared local
spans overlap. **After**, reading member *X* requires that *X*'s own local
layout validate and that no other record's declared local span intersect *X*'s
span. Nothing else changed.

Concretely, this is what is no longer checked for a record the caller is not
reading and whose declared span does not reach the target:

- its local and central compression method, flags, name, sizes and CRC;
- its local ZIP64 size framing;
- its data descriptor's contents;
- whether its span overlaps a *third* record that the target also does not
  touch;
- whether its span runs into the central directory.

What is still archive-wide: **distinct local-header offsets**, at zero I/O.

Change 0575's witness makes the delta exact. It is a 4,405-byte archive of three
Store members, each individually self-consistent — local and central names,
methods, flags, CRCs and sizes all agree for all three:

```
member    local_off   csize  local_extra  declared span
A.bin             0      32         4096   [0, 4163)      <- contains B's whole local record
B.bin          2048      32            0   [2048, 2115)
C.bin          4163      32            0   [4163, 4230)
central directory at 4230
```

`A.bin`'s **local** extra field is 4,096 bytes; its **central** record declares
an extra length of 0. The central directory therefore cannot see that A's span
swallows B — only the two bytes at `offset 28..30` of A's local header reveal
it.

| target | before | after |
| --- | --- | --- |
| `A.bin` | REFUSE | **REFUSE** |
| `B.bin` | REFUSE | **REFUSE** |
| `C.bin` | REFUSE | **accept**, returning its 32 bytes |

**Corrected after measurement: that one row is _not_ the entire semantic
change, and an earlier draft of this record said it was.** The witness
demonstrates the *overlap* case, which is the case change 0575 designed around
and the case the project owner was shown when approving this change. Change
[0582](0582-zip-strict-scope-differential-fuzz.md) then measured the delta over
22,875 inputs and found it spans **twelve refusal identities, not one**: the
overlap row is 37,176 member verdicts, and a further 1,094,221 verdicts come
from every *other* local-versus-central consistency check on a record the caller
does not read — sizes, flags, CRC, method, names, ZIP64 framing, `Eof`, a
missing local signature, and disk-start metadata. The approved row is **3.3%** of
the whole.

The bulleted list earlier in this section was accurate in the original draft and
is unchanged; what was wrong was the sentence claiming the witness exhausted it.
The accurate one-sentence statement of this change is: **every
local-versus-central consistency check on every record the caller does not read
is dropped**, and overlap is one instance of that. The owner re-approved on the
corrected, measured delta after change 0582 reported it; that re-approval, not
the original one, is what authorises this change. Change 0582 also measured that
**no archive in `test-data/` changes behaviour at all** — 516 real containers and
7 seeds, zero divergences of any class — so the delta lives entirely in malformed
and adversarial inputs.

Within the overlap case the witness illustrates, an archive in which two members
the caller never touches overlap each other is now read successfully for every
member that is not itself overlapped. Both overlapping members are still
refused, from either side: reading `A.bin` is refused because `B.bin` is a
successor whose offset lies before A's exact span end; reading `B.bin` is
refused because A is inside B's residual window and A's 30-byte header reveals
`span_end_A = 4163 > 2048`.

The realistic OOXML shape: a producer emits `xl/media/image1.png` with an
inflated local extra field so its span swallows `xl/media/image2.png`. Before,
reading `xl/worksheets/sheet1.xml` failed. Now the sheet reads and either image
is refused.

The same contract applies to both readers. `ArchiveReader` (slice-backed) pays
no positional read for any of this, so the change buys it CPU rather than I/O,
and it *loses* an archive-wide property it had for free. That was change 0575's
explicit recommendation and its stated reason: one wording beats a library where
the same bytes are accepted or refused depending on whether the caller supplied
a path or a `Vec<u8>`.

One path is deliberately **not** narrowed. `ArchiveReader::read_stored_borrowed`
reaches `validate_borrowed_spans`, a separate archive-wide loop with its own
error strings (`"borrowed access cannot prove non-overlapping ZIP spans"`), and
it publishes a borrowed slice into caller hands under a contract change 0575
does not analyse. It keeps its archive-wide proof unchanged.

## Re-deriving the residual window: 0575's bracket is not sound for (b)

Change 0575 computes a predecessor's zero-I/O bracket as

```
min_end_i = offset_i + 30 + central_name_len_i + compressed_size_i
max_end_i = min_end_i + 65535 + 24        // "the residual window is 65,559 bytes"
```

and justifies the single `u16` with: "The strict path forces `local_name_len ==
central_name_len` (a differing local name is refused outright)".

That is true of the record being **read**. It is not true of a record candidate
(b) never validates. A predecessor may declare a local `file_name_length` its
central record does not carry — its own read would be refused with `"strict
local and central names differ"`, but nobody is reading it — and both halves of
the local variable region are `u16`. The sound bracket is therefore

```
min_end_i = offset_i + 30 + compressed_size_i
max_end_i = offset_i + 30 + 2*65535 + compressed_size_i + 24     // 131,094 residual
```

The attack the 65,559-byte window admits is concrete and is now a test. A
predecessor at offset 0 with a 9-byte central name and a 32-byte payload has a
0575 bracket ending at `30 + 9 + 32 + 65535 + 24 = 65,630`. Give it a local
header declaring `file_name_length = 100` and `extra_field_length = 65535` and
its declared span ends at `30 + 100 + 65535 + 32 = 65,697`. A member at offset
65,640 lies inside that span and outside that bracket: the 65,559-byte window
prunes the record that claims its bytes. The wider window reads the header and
refuses. `a_predecessor_reaching_past_the_central_name_length_bracket_is_still_refused`
pins it.

The correction doubles the window, and that is where 0575's headline prediction
breaks. It costs reads; it never admits a read the narrower window would have
refused, so it cannot widen the semantic delta — only narrow it.

## Measured effect

Positional reads and bytes issued by the strict-layout path alone, counted
through a `ReaderAt` that records every `read_at`, on the source-backed
`IndexedArchive`. Both trees are `git archive` extractions of the committed
revision `32d25e088` with their own `CARGO_TARGET_DIR`, so a concurrent edit
under `crates/` cannot reach them; the after tree is that extraction with this
change's two files overlaid. "before" is the **post-0573** tree, not the
pre-0573 one.

| fixture | scenario | pre-0573 | post-0573 (before) | after | after bytes |
| --- | --- | ---: | ---: | ---: | ---: |
| `sheet-names.xlsx` (13) | open, list | 0 | 0 | **0** | 0 |
| | one member | 26 | 13 | **8** | 2,478 → **264** |
| | one-cell closure (6 parts) | 26 | 13 | **11** | 2,478 → **1,742** |
| | all 13, physical order | 26 | 13 | **13** | 2,478 → 2,478 |
| | all 13, reverse order | 26 | 13 | **25** | 2,478 → 2,838 |
| `ConditionalFormattingSamples.xlsx` (132) | open, list | 0 | 0 | **0** | 0 |
| | one member | 264 | 132 | **20** | 10,740 → **624** |
| | one-cell closure (7 parts) | 264 | 132 | **59** | 10,740 → **3,202** |
| | all 132, physical order | 264 | 132 | **132** | 10,740 → 10,740 |
| | all 132, reverse order | 264 | 132 | **263** | 10,740 → 14,670 |
| `shapes.pptx` (48) | open, list | 0 | 0 | **0** | 0 |
| | one member | 96 | 48 | **15** | 4,748 → **471** |
| | one-slide closure (6 parts) | 96 | 48 | **16** | 4,748 → **1,918** |
| | all 48, physical order | 96 | 48 | **48** | 4,748 → 4,748 |
| | all 48, reverse order | 96 | 48 | **95** | 4,748 → 6,158 |

The pre-0573 column is change 0573's retained measurement, not remeasured here.
The post-0573 column was recaptured for this change and reproduces 0573's
figures exactly — 13 / 2,478, 132 / 10,740, 48 / 4,748 — which is the control
that the two changes are separable by region.

Change 0575's corpus is three descriptor-free fixtures from one producer family,
and its own Limitations say so. Three descriptor-bearing fixtures were added
because they are the adverse case:

| fixture | members | descriptors | one member | closure | all, order | all, reverse |
| --- | ---: | ---: | --- | --- | --- | --- |
| `xlsx/universal-content.xlsx` | 13 | 13 | 39 → **9** | 39 → **36** | 39 → 39 | 39 → **63** |
| `docx/comment.docx` | 10 | 10 | 30 → **11** | 30 → **11** | 30 → 30 | 30 → **48** |
| `pptx/shape-soft-edges.pptx` | 65 | 65 | 195 → **65** | 195 → **86** | 195 → 195 | 195 → **323** |

### Convergence, over the whole OOXML corpus

Reading **every** member of every DOCX/PPTX/XLSX fixture under `test-data/ooxml`
in physical order, before and after:

```
fixtures 168   members 4089
reads  4385 -> 4385      bytes 517371 -> 517371      accepted 4089 -> 4089
```

The per-fixture accept/refuse verdict string, read count and byte count are
**byte-for-byte identical** across all 168 fixtures (`diff census-before.txt
census-after.txt` is empty). Reading the whole archive converges on exactly the
archive-wide proof, at exactly the archive-wide cost.

### The regression, stated plainly

**Reading every member in reverse physical order costs up to 2n−1 reads where
the archive-wide proof cost n.** A record probed as a neighbour (30 bytes) and
later read as a target needs its full local header, and that is a second read of
the same record. Reading forward, a validated target is its own exact span bound
for everything after it, so no record is read twice; reading backwards, nearly
every record is. Measured: 13 → 25, 132 → 263, 48 → 95, 65 → 323 across the six
fixtures above, and bytes rise between 15% (`sheet-names.xlsx`, 2,478 → 2,838)
and 57% (`comment.docx`, 730 → 1,144).

Verdicts are unaffected — the same six read orders produce the same three
verdicts on the witness, and `a_verdict_does_not_depend_on_what_the_reader_read_before`
pins that. This is a cache-cost property, which ADR 0005 permits to vary, but it
falsifies change 0575's claim that (b) "is never more expensive than today", and
it is recorded rather than worked around. Retaining each probed record's 640-byte
header window instead of its 30 bytes would remove it at a retention and
bytes-moved cost; that trade was not taken.

## Scoring against change 0575's frozen prediction

| fixture | scenario | 0575 predicted | measured | verdict |
| --- | --- | ---: | ---: | --- |
| `sheet-names.xlsx` | open, list | 0 | 0 | **confirmed** |
| | read one member | 8 | 8 | **confirmed** |
| | one-cell closure | 11 | 11 | **confirmed** |
| | read all 13 | 13 | 13 | **confirmed** |
| `ConditionalFormattingSamples.xlsx` | open, list | 0 | 0 | **confirmed** |
| | read one member | **4** | **20** | **falsified**, 5× |
| | one-cell closure | **14** | **59** | **falsified**, 4.2× |
| | read all 132 | 132 | 132 | **confirmed** |
| `shapes.pptx` | open, list | 0 | 0 | **confirmed** |
| | read one member | 15 | 15 | **confirmed** |
| | one-slide closure | **15** | **16** | **falsified**, +1 |
| | read all 48 | 48 | 48 | **confirmed** |

Nine of twelve confirmed, three falsified, for **two independent reasons**:

1. **The residual window.** The 132-record fixture is the only one of the three
   large enough for the window size to matter — the whole of `sheet-names.xlsx`
   and `shapes.pptx` fits inside either window. Re-deriving the bracket soundly
   doubles it and takes one member from 4 reads to 20 and the closure from 14 to
   59. `results/change-0580/residual_window.py` prices both brackets and
   reproduces both columns.
2. **Ordering.** 0575's cost model counts the *union* of records each target
   touches. The implementation memoises per record, so a record first probed as
   a neighbour and then read as a target costs two reads, and the union
   under-counts by exactly that number. On `shapes.pptx`'s closure,
   `ppt/slides/_rels/slide1.xml.rels` is probed while proving
   `ppt/_rels/presentation.xml.rels` and read afterwards: 15 → 16. The same
   script, with the memo modelled in read order, reproduces all six measured
   scenario figures exactly (8, 11; 20, 59; 15, 16).

Byte predictions: `sheet-names.xlsx` 264 / 1,742 / 2,478 — **all three exact**.
`shapes.pptx` 471 / 1,888 / 4,748 — 471 and 4,748 exact, 1,888 measured as 1,918
(one extra 30-byte probe, the same ordering effect).
`ConditionalFormattingSamples.xlsx` 144 / 1,852 / 10,740 — 10,740 exact, the
other two falsified with the read counts (624 and 3,202).

The four predictions 0575 made independently of the counts:

| prediction | outcome |
| --- | --- |
| Open and list stay at **zero** strict-layout reads on all three fixtures | **confirmed** |
| `read`/`read_entry` stay at zero, before and after, because they do not reach the proof | **confirmed**; `opening_and_listing_issues_no_strict_layout_read` asserts `read_entry` does not even prime the cache |
| Reading every member converges to exactly the change-0573 count, no record's header read twice | **confirmed in physical order** over 168 fixtures; **falsified as an order-independent statement** — see the regression above |
| The witness is refused for `A.bin` and `B.bin` and accepted for `C.bin`, and the three named overlap tests keep their verdicts and their `ErrorKind::InvalidInput` | **confirmed**; all 454 pre-existing `soapberry-zip` tests pass unchanged |

Change 0575's falsification criterion 5 — "a residual set proves unbounded in
practice" — is partly realised and worth naming: the worst single target on
`shapes.pptx` still reaches all 47 predecessors, and on
`ConditionalFormattingSamples.xlsx` the sound bracket raises the worst target
from 39 to 43 of 131. The scenario-level figures above are what the criterion
was about, and they remain well below *n*.

## Validation preserved

- **`validate_reader_entry_layout_with_name_policy` is untouched**, byte for
  byte. The target keeps change 0573's check order exactly: metadata →
  `local_header_offset > central_directory_offset` → fixed-read I/O → signature
  → `local_end > central_directory_offset` → variable-read I/O → method → flags
  → name → size framing → sizes → CRC → data span → descriptor → `span_end`.
- **Change 0573's payload-avoidance property is intact.** The target still
  receives the following record's local offset as its read hint, and that hint
  is still sound: the successor check proves `span_end_target <=
  offset_of_next_record`, which is exactly the inequality 0573's window bound
  relies on. A neighbour probe reads **exactly 30 bytes** at an offset the
  central directory itself declares to be a local header, and never a variable
  region or a payload byte. The only other neighbour read is a data descriptor
  at that record's declared payload end, which is the identical read the
  archive-wide proof already issued for that record, bounded by the
  central-directory offset by `DataDescriptor::parse_complete_at`.
- **A data descriptor's width is resolved exactly, never reserved.** A
  descriptor's encoded width is 12, 16, 20 or 24 bytes and is not decidable from
  a fixed local header. In a gapless archive a predecessor's payload ends
  exactly one descriptor before the next record, so a bound that simply reserved
  the widest form would refuse **every** gapless descriptor-bearing archive —
  three of the six measured fixtures, 148 members of the corpus. When and only
  when the reserved range decides the verdict, the width is settled with one
  read using the same hint `resolve_local_entry_size_framing` computes.
  `a_descriptor_bearing_predecessor_is_resolved_exactly_not_conservatively`
  pins it for signed and unsigned descriptors.
- **Duplicate local-header offsets are still refused**, now with their own error
  and at zero I/O, over every record including directory records. Change 0575
  found this branch had no test and was in fact unreachable — an archive with
  equal offsets tripped a name mismatch first. It is reachable and tested now.
- **Order independence.** A verdict is a function of the archive bytes and the
  target. The memo records successes only and stores facts, never verdicts; a
  failed proof restores the memo it started from, so a read that failed for
  cancellation, a budget error or a transient source error retries in full. This
  is what disqualified change 0575's candidate (c) under accepted ADR 0005
  ("Cache behavior is semantically invisible"), and a six-order test pins it.
- **The single-flight contract is unchanged.** The memo is moved out of the
  cache for the duration of one proof, so a second thread still waits on the
  condvar and a source that re-enters its own archive on the proving thread
  still reports `"strict layout proof build re-entered on its owning thread"`
  instead of deadlocking. `indexed_strict_first_read_uses_one_single_flight_builder`
  and `indexed_strict_streams_reuse_the_bounded_layout_path` pass unchanged.
  Change 0575's concurrency goal — disjoint members not serialising — is
  therefore **not** delivered; see Limitations.
- **Bounded resources.** Every new allocation is a `try_reserve` /
  `try_reserve_exact` on a named resource. Retention falls: the proof no longer
  retains a `StrictEntryLayout` and a hash entry for all *n* records, only for
  records actually touched, plus one `u64` per record of prefix maximum (1,056
  bytes at *n* = 132). The predecessor scan is `O(n)` arithmetic in the worst
  case and stops at the prefix maximum in the common one.
- No `unsafe`; `deny(unsafe_code)` intact. Preservation
  (`validate_preservation_entry_layout`) is untouched and still shares the
  per-entry prover with `allow_name_mismatch: true`.

## ADR compliance

Change 0575's matrix assessed candidate (b) against ADRs 0001, 0005, 0006 and
0010/0011 and recorded **no ADR exception requested**, gated on a range-source
measurement. That gate was satisfied by change
[0572](0572-ooxml-range-source-attribution.md), which measured the proof as 264
of the 354 requests an XLSX cell read issues on a caller-supplied source. The
implementation does not move the assessment:

- **ADR 0001** — the refusal stays a typed `Result` and still precedes any sink
  write; no API surface moves and no implementation type is exposed.
- **ADR 0005** — strictly less I/O on every selective scenario, and cache
  behaviour stays semantically invisible: the memo holds facts about the
  archive's bytes, never verdicts, and holds successes only. The one deviation
  worth naming is that cache behaviour is now visible in *cost* in a way it was
  not: reverse-order traversal costs up to 2n−1 reads. ADR 0005 constrains
  semantics, not cost, but the regression is reported above rather than buried.
- **ADR 0006** — every byte returned is proven not to lie inside another
  record's declared local span. The delta is confined to records the caller
  never reads, and invariants (3) and (4) of change 0575's numbering are kept in
  full at zero cost. Preservation, output determinism and every limit are
  untouched.
- **ADR 0010 / 0011** — both readers stay in `soapberry-zip`, the ZIP grammar
  owner below `litchi-opc`. Nothing crosses a crate edge.

## Correctness evidence

Eleven tests were added. **Six fail against the pre-change code**, verified by
running them against a `git archive` extraction of `32d25e088` with the two
prover helpers inlined:

| fails pre-change | what it pins | pre-change outcome |
| --- | --- | --- |
| `target_scoped_layout_refuses_both_overlapping_members_and_admits_the_third` | change 0575's witness: A and B refused, C accepted, on **both** readers | `C.bin` refused with `"…refuses overlapping ZIP local spans"` |
| `a_verdict_does_not_depend_on_what_the_reader_read_before` | six read orders on one reader and on fresh readers, identical verdicts, both readers agreeing member by member | C's verdict is `false` in every order |
| `duplicate_local_header_offsets_are_refused_on_the_strict_path` | the branch change 0575 found untested — and unreachable | reported `"strict local and central names differ"` |
| `a_predecessor_reaching_past_the_central_name_length_bracket_is_still_refused` | the sound residual window: a predecessor with an inflated local **name** length, 65,640 bytes away, still refused | reported `"strict local and central names differ"` |
| `a_predecessor_outside_the_residual_window_costs_no_read` | six tiny members before a 140,000-byte one: reading the target costs 2 reads, not 8 | 8 reads |
| `a_failed_proof_memoises_nothing_and_the_next_read_retries` | a refusal for one member neither poisons nor primes the reader | reading `C.bin` after `A.bin` fails |

Five pass on both sides, which is the point of them — they are the preservation
half of the contract:

| passes on both sides | what it pins |
| --- | --- |
| `a_distant_predecessor_that_reaches_the_target_is_still_refused` | a 16 KiB inflated local extra field covering three later members: all three refused, on both readers. This is the case the central-directory bracket cannot decide alone and the whole reason predecessor headers are read. |
| `a_descriptor_bearing_predecessor_is_resolved_exactly_not_conservatively` | signed and unsigned descriptors, three-member archives, every member readable |
| `zip64_and_descriptor_members_keep_their_target_scoped_proof` | Store and Deflate members written by the crate's own writer, plus descriptor-bearing archives, on both readers |
| `opening_and_listing_issues_no_strict_layout_read` | the control: open and list issue **zero** reads, and `read_entry` does not reach the proof at all |
| `reading_every_member_converges_on_the_archive_wide_verdict` | six archive shapes: the target-scoped verdict of every member agrees with an archive-wide verdict computed in-test, and reading every member in physical order costs exactly what one archive-wide proof cost |

The witness archive is generated, not committed as a blob: `overlap_witness_archive()`
builds the same 4,405 bytes `results/change-0575/overlap_witness.py --out`
writes, including the `0xFACE` unknown extra-field header, and asserts the size.

Corpus-scale convergence is evidence rather than a test, because the
`soapberry-zip` suite is deliberately self-contained and reads no fixture from
`test-data`: `results/change-0580/census-before.txt` and `census-after.txt` carry
a per-member accept/refuse string for all 168 fixtures and 4,089 members, and
they are identical.

Gates run for this change:

- `cargo fmt -p soapberry-zip -- --check` clean;
- `cargo clippy -p soapberry-zip --all-targets` clean, and
  `cargo clippy -p soapberry-zip --all-features --lib --no-deps -- -D warnings`
  clean (the crate does not opt into the workspace lint table, so both runs are
  explicit);
- `RUSTDOCFLAGS="-D warnings" cargo doc -p soapberry-zip --no-deps` clean;
- `cargo check -p soapberry-zip --no-default-features` clean;
- `soapberry-zip` **465 lib tests + 116 integration and doc tests, zero
  failures** (454 lib tests before, 11 added);
- `litchi-opc` **664 tests, zero failures**;
- the OOXML facade suites `litchi-xlsx`, `litchi-docx` and `litchi-pptx`
  together, **3,597 tests, zero failures**.

The `parse_zip` fuzz target was **not** run: `cargo-fuzz` is not installed in
this environment and no nightly toolchain is present. That gate is outstanding
and is recorded as such rather than claimed. Change 0575's admission list names
it; it is the one item on that list not discharged here.

## Limitations

No timing, allocation-profile, cold-cache, physical-device, cross-platform or
producer-corpus result is claimed. The counts are logical `ReaderAt` calls over
warm fixtures. As change 0573 put it, a reduction in positional reads on a warm
local file is not separable from timing noise; change 0572 is the record that
prices these requests against a latency-bearing source, and nothing here asserts
a wall-clock result.

Five gaps are recorded rather than worked around.

**Reverse-order traversal costs up to 2n−1 reads.** Measured and quantified
above. Verdicts are unaffected; cost is.

**The concurrency win in change 0575's design was not taken.** 0575's retention
table promises that "disjoint members proceed independently" under (b). This
implementation keeps the whole-proof single flight instead, because that is what
preserves the re-entrancy error, the `build_count`/`is_ready` hooks and the two
existing concurrency tests verbatim. Two threads reading different members of
the same archive still serialise, exactly as before. Lifting that is a separate
change with its own contract questions.

**The residual window is a worst case, not a distribution.** 131,094 bytes is
the sound bound over every archive a `u16` can describe. On the measured OOXML
family it prunes nothing on a 9 KB or 68 KB archive, because those archives fit
entirely inside it, and prunes 111 of 131 predecessors on the 654 KB one. An
archive whose members are large prunes more; an archive of thousands of tiny
members prunes less, and its worst single target can approach *n*, at which
point (b) costs what the archive-wide proof cost while proving strictly less.
Change 0575's falsification criterion 5 is the standing rule for that case.

**The predecessor scan is `O(n)` arithmetic per target in the worst case.** The
prefix-maximum array stops it early on every measured fixture, but an archive
whose records all declare payloads reaching toward one target defeats the stop.
Reading every member of such an archive is `O(n²)` comparisons where the
archive-wide proof was `O(n)`. No additional I/O is involved — each record's
header is still read at most once per reader — and *n* is bounded by the
existing `max_metadata_bytes` and declared-entry-count limits, but the
amplification is real and unmeasured.

**Resolving a neighbour's data descriptor compares it against that neighbour's
central record.** There is no way to learn a descriptor's encoded width without
disambiguating its framing, and disambiguation is the comparison. So a
descriptor-bearing predecessor whose span end decides the target's verdict does
get one check applied that pure target scoping would not require, and a
neighbour with a corrupt descriptor in that position refuses the target. This is
narrower than change 0575's design, never wider, and it affects only a
predecessor within 24 bytes of the target's local header.

Finally, the archaeology in change 0575 — no accepted ADR, no threat model, no
rationale in either introducing commit — is a negative finding over this
repository at that revision, and this change inherits it. If a rationale for the
archive-wide scope exists outside the repository, 0575's falsification criterion
1 applies and this change is the thing to revisit.
