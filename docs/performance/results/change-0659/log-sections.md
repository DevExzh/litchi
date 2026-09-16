# Log paragraphs for change 0659

Four blocks, one per merged log. Each is written to be prepended as the newest
section of its file.

## For `HOTSPOTS.md`

## 0659 — the DOC open's six complete scans become three, and the 84% of the corpus that is refused pays for it too

Change 0644 enumerated the DOC snapshot open read by read and froze three
variants; decision 9 of change 0652 authorized the most aggressive one that does
not widen the trailing window, **B2 + C**, and this change lands it. The open now
takes **three** complete artifact scans instead of six and **four** CFB index
parses instead of five: 30 reads become 23 on `documentProperties.doc` and 46
become 34 on `picture.doc`, with `open` plus `paragraph(0)` falling from 59 to
50 and from 14 complete scans to 9. The PPT open is **2 → 2** and byte-identical
in every read on all 30 fixtures, because variant B2's rule — one scan for every
empty-splice identity call *except the last of its operation* — exempts an open
that makes only one. Natively that is a median **−35.4%** of the DOC open's
cycles (−34.90% to −49.45% over the 8 admitted fixtures), a median **−25.0%** of
`open` plus `paragraph(0)`, and **−0.12%** on PPT, against an in-window A/A floor
of +0.06% median and 0.62% at p95; paired wall-clock timing agrees at
−34.3%…−49.5% and −23.5%…−32.3% at p50 with a 0.71% p95 floor. The saving is
**larger than change 0644 modelled**, and the reason is instructive rather than
alarming: 0644 priced only the per-byte hashing term. Fitting the measured saving
as `3 × c × bytes + fixed` recovers `c = 2.1471` cycles per byte per scan —
0644's own constant to 0.75% over a 149× size range, so exactly three scans came
out — plus a **53,780-cycle fixed term** that is the composed CFB reopen the
third identity call took with it and the 45 metadata observations (35 `version()`
and 10 `len()` per open on `documentProperties.doc`) that went with the removed
scans. The readback's fit separates the two at about **9,608 cycles per removed
`fingerprints` call** and **24,956 for the removed reopen**. The result 0644
predicted would be invisible in the open median is the largest one on the corpus:
49 of 57 `.doc` fixtures are refused inside the semantic parse and never reach
identity call 2 or 3, so Option C does nothing for them and variant B2 is the
whole of their **−42.3% median in bytes read**; the 8 admitted fixtures fall from
6.05×–7.47× the artifact to 3.04×–4.21×.
[Change and limitations](0659-cfb-single-scan-identity-entry-point.md);
[evidence](results/change-0659/README.md).

## For `GOAL_AUDIT.md`

## 0659 — a fence reduced where the caller already brackets it, and an error precedence repaid rather than dropped

`docs/GOAL.md` puts correctness and lossless preservation over speed, and this
change is the first in the programme to *remove* a complete-artifact scan from an
ADR 0006 identity fence. What makes it admissible is stated as a property of the
call site, not of the primitive: a reduced identity call is one whose comparison
partner is a scan the same operation already takes — the open's first call is
compared against identity call 2's planning scan, and `ensure_current`'s first is
compared twice, against the digest retained at open and against `confirmed`. The
PPT open has no such partner and is therefore untouched, obtained by applying
B2's rule rather than by special-casing PPT. Change 0644's refusal of Option A
stands unamended and its witness was rerun: a `SourceVersion` observation cannot
replace any scan, because an in-place `pwrite` with the modification time
restored leaves both `version()` and `len()` unchanged. Two contract movements
are accepted and both are written into the crate documentation rather than left
in a record. Option C ends the open's fence one identity call earlier, so six
read ordinals that were refused by `open` now return `Ok` and are refused by the
first access instead — with `Overlay(SourceFingerprintChanged)` on a payload
witness and `Overlay(Ole(CorruptedFile))` on a FAT-sector one, which is exactly
what 0644's gate said to expect — and the five accessors that consume no byte
(`source_version`, `fingerprint`, `len`, `is_empty`, `limits`) now say in their
rustdoc that they return state retained at open. The second movement is the one
0644 avoided by recommending a weaker variant, and the audit should note how it
was repaid: removing the open's first confirming scan hands the window in which
a caller's own writer corrupts the artifact's allocation table to the retained
index parse, which would report `CorruptedFile` and blame the file. Change
0621's relocation rule is applied there — narrowly, to that parse's error alone,
to the four `OleError` variants that describe structure, and only on positive
proof that the digest moved — using a new reopen-free digest entry point,
because every entry point that reopens the composed candidate fails on the same
damage instead. The relocated error is the *same value*, not merely the same
variant: the sweep's rendering after the change is character-for-character the
before leg's, including the expected digest. One surviving ordinal does change
its typed error, and the record says so plainly rather than absorbing it: the
confirming scan was the only observation that distinguished "the mutation
arrived before the retained parse" from "it arrived after the identity was
confirmed", so with it gone both are reported as what they are, a change of the
source. Variant B1 is **refused** on the evidence 0652 asked for: ordinals 45–49
of the landed `open` plus `paragraph(0)` are refused today and would fall outside
B1's last scan, so its wider trailing window does admit mutations B2 catches. No
limit was relaxed, no defence removed and no ADR amended. OLE2 and OOXML remain
the active priority; ODF stays deferred and iWork excluded.
[Change and limitations](0659-cfb-single-scan-identity-entry-point.md);
[evidence](results/change-0659/README.md).

## For `REPORT.md`

## 0659 — B2 + C landed, with both public entry points change 0644 priced

Two public items are added to `litchi-cfb`'s `SharedOleFile`:
`caller_bracketed_identity_fingerprint`, the empty-splice identity capture that
takes one complete scan and keeps the composed CFB reopen, whose rustdoc states
that the caller owns the read-twice-compare and when the method may not be used;
and `unbracketed_source_fingerprint`, a digest-only path with no reopen whose
rustdoc states that it is a failure-path diagnostic and never an identity
capture. The first is the empty-splice specialization of the existing
`finish_overlay_plan_with_owner`, factored into one private helper so it cannot
drift from it; the second calls the private `fingerprints` directly, and the
`SourceSnapshot` both construct becomes one accessor shared with the three
planners that used to write it out inline. `litchi-doc` applies the first to its
three identity calls that
are not the last of their operation, deletes the open's third identity call while
keeping its `ensure_source_identity`, and gives the retained index parse an error
branch that uses the second. `litchi-ppt` is unchanged. Evidence is deterministic
first: complete scans **6 → 3** for the DOC open on all 8 admitted fixtures,
**14 → 9** for `open` plus `paragraph(0)` where the read is admitted and
**10 → 6** where it is refused, **2 → 2** for the PPT open on all 30 — the
counts change 0644's gate G3 requires for the declared variant, and the
`version()`/`len()` reductions (108/32 → 73/22 on `documentProperties.doc`,
132/32 → 87/22 on `picture.doc`) are reconciled term by term against the removed
`fingerprints` calls and the removed composed reopen with nothing left over, the
five `ensure_source_identity` sites untouched. A held mutation was placed after
**every** read ordinal of a clean run on three witness placements — outside every
parsed range, inside the first FAT sector, inside the first directory sector —
over 8 `.doc` in two modes and all 30 `.ppt`, on both legs: 144 sweeps and 3,444
mutated opens before, 3,009 after. No interior ordinal became `OK`, the leading
window is identical everywhere, the trailing window is one ordinal on both legs
of every DOC sweep, and all **92** PPT sweep files have the same sha256 on both
legs. Change 0589's 87-artifact publication differential is **byte-identical**.
Native `perf stat -r 3` isolation pairs in the order A1 B1 B2 A2 give a DOC open
median of **−35.41%** (−34.90%…−49.45%), a readback median of −24.96%, and PPT
−0.12% with not one fixture adverse by more than 0.16%, against an A/A floor of
+0.06% median and 0.62% at p95 over 76 same-leg pairs. Paired timing over 120
samples a leg agrees: DOC open p50 −34.34%…−49.49% (before against after
+52.30%…+97.96%), readback −23.45%…−32.27%, PPT +0.12% and +0.21% — the only two
scenarios that got worse, both inside the 0.71% p95 floor. Gate G4's literal ±5% band against change 0644's predicted after-value
**fails** on 7 of 8 DOC fixtures and the record says so: the prediction omitted
the fixed term, the count proves no further scan was removed, and the
decomposition fits the whole corpus to ±0.7%. No claim-registry entry;
`performance_claim: none`.
[Change and limitations](0659-cfb-single-scan-identity-entry-point.md);
[evidence](results/change-0659/README.md).

## For `ADR_COMPLIANCE.md`

## 0659 — the first reduction of an ADR 0006 fence, admitted by the caller's bracket and by nothing else

ADR 0005 requires that a mutation during a read return `SourceChanged`, ADR 0006
binds a splice to the bytes it was planned against, and ADR 0003 makes the
fingerprint diagnostic while exact byte equality authorizes application. All
three are preserved, and the argument is deliberately narrow. The removed
confirming scans are admissible **only** because each reduced call's comparison
partner is a scan the same operation already takes; the retained
`ensure_source_identity` (a `version()`, a `len()` and a `version()`) carries no
part of the argument, stays at all five fence points of the open and all three
inside `ensure_current`, and is proved insufficient by change 0644's `FileSource`
witness, rerun here identically. The new entry point's rustdoc states the
condition as a contract — a caller with no outer comparison must keep using
`plan_same_length_stream_splices` — and the PPT open, which is such a caller, is
untouched and measured to be byte-identical read for read on all 30 fixtures.
ADR 0003's composed reopen is retained by both the reduced and the paired call,
so Option D stays rejected and 14 of the FAT sweep's 17 `Overlay(Ole(..))` rows
keep both wrapper and site. ADR 0005's `SourceChanged` rule holds at every
ordinal the sweeps cover, with the two movements the record enumerates: Option
C's six relocated ordinals, refused at the first access with the same variant on
the payload witness and with the composed reopen's own `Overlay(Ole(..))` on the
FAT witness; and one re-typed ordinal where a structural CFB failure of the
retained index parse over moved bytes is now attributed to the movement rather
than to the file. That second movement is change 0621's relocation rule, scoped
to that parse's error, to the four `OleError` variants describing structure —
a limit refusal, an I/O failure, an allocation failure and an already-typed
`SourceChanged` keep their own error exactly — and conditioned on a recomputed
digest actually differing, so it cannot fire on a file that was malformed before
anything was read, and it cannot fire on the composed reopen's own error, which
is the inversion change 0644's gate fails an implementation for. No limit was
relaxed, no `unsafe` added, no dependency added, no archive type, raw lock or
executor leaked, and the facade is untouched. No ADR text is amended, because
decision 9 of change 0652 assigns this change none; ADRs 0030 and 0031, now
Accepted, are not cited because nothing here is lazy or budgeted.
[Change 0659](0659-cfb-single-scan-identity-entry-point.md);
`performance_claim: none`.
