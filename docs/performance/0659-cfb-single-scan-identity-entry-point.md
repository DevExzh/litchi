# 0659: the DOC snapshot open goes from six complete scans to three, and the confirming scan it gives up is repaid by a reopen-free re-hash that keeps the one error whose name would otherwise change

Status: **retained, implemented.** `performance_claim: none` — the read counts,
the change-under-read sweeps, the witnesses, the native cycles and the paired
timing below are reported as evidence, not registered as a claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements variant **B2 + C** of the design change
[0644](0644-ole2-snapshot-fence-design.md) froze, on the authority of decision 9
of change [0652](0652-owner-decisions-for-the-third-wave.md). It is queue row 9
of [0651](0651-queue-refresh-after-the-second-wave.md).

## What was changed

Three edits, in two crates.

1. **`litchi-cfb` gains two public entry points** on `SharedOleFile`
   (`crates/litchi-cfb/src/overlay.rs`):
   `caller_bracketed_identity_fingerprint`, the empty-splice identity capture
   with its confirming scan removed and its composed CFB reopen kept; and
   `unbracketed_source_fingerprint`, a digest-only failure-path diagnostic with
   no reopen at all. The first is the empty-splice specialization of the
   existing `finish_overlay_plan_with_owner`, factored into one private helper
   (`empty_splice_identity`) so it cannot drift from it; the second is a direct
   call to the private `fingerprints`. The `SourceSnapshot` both construct was
   written out inline at three existing sites, so it becomes one accessor
   (`plan_source_snapshot`) rather than a fourth copy.
2. **`litchi-doc`'s source-backed snapshot takes variant B2**
   (`crates/litchi-doc/src/body_text/source.rs`): the three identity calls that
   are not the last of their operation — the generic open's first
   (`:531`), the owned open's first (`:686`) and `ensure_current`'s first
   (`:821`) — take one complete scan instead of two. The retained index parse
   at `:543` grew an error branch, `relocate_structural_cfb_error`, which is
   change 0621's relocation rule applied to cost B-iii.
3. **`litchi-doc` takes Option C**: the open's third identity call is deleted.
   Its `ensure_source_identity` fence point is kept, so the open still has the
   same five.

**`litchi-ppt` was not edited.** Variant B2's rule — *B applies to every
empty-splice identity call except the last one of its operation* — exempts the
PPT open, whose single identity call is the last of its operation, and PPT's
`ensure_current` (`text_edit.rs:515`) makes no identity call at all. That is
0644's refusal of Option B on PPT, obtained by applying the rule rather than by
special-casing it, and the gates below measure that PPT did not move.

The resulting DOC open is

    I1 → R1 → I3a → I2 → semantic parse → R3 → I3b → R4 → compare(R1, R3)

— **three complete scans** from six, **four index parses** from five, the same
five `ensure_source_identity` points, and a trailing unbracketed window of one
ordinal, as before.

## Authority

Decision 9 of change 0652, in the owner's words:

> "DOC snapshot's source-identity fence: add a new public entrypoint for it."

0652 reads that as authorizing "the second public `litchi-cfb` entry point
0644's variant B2 needs, and B2 + C (the DOC open from six complete scans to
three); B1 only if its wider trailing window is shown by witness not to admit a
mutation B2 catches", and requires this record to prove "0644's four admission
gates; every witness in 0644's sweep re-run on the landed change with the same
verdicts; the error identity 0644 mapped preserved by name".

Standing trade-off 1 of 0652 ("breaking changes are totally acceptable")
authorizes the public additions and the documented movement of the retained
accessors' contract. Standing trade-off 2 ("correctness and safety is the
primary consideration") is what decides section *B1 is not taken* below.

### B1 is not taken, and the witness says why

0652 admits B1 only on a witness that its wider trailing window "admits no
mutation B2 catches". The landed B2 sweeps exhibit the opposite, without
building B1: on `open` plus `paragraph(0)` over `documentProperties.doc` the
operation now takes 50 reads, and ordinals **45–49** — the position of the
trailing `ensure_current`'s last confirming scan, plus the four reads of the
composed reopen that precedes it — are refused today with
`read-ERR Overlay(SourceFingerprintChanged)`
([`sweeps/sweep-documentProperties-doc-open-read-after.tsv`](results/change-0659/sweeps/sweep-documentProperties-doc-open-read-after.tsv),
band 23–49). B1 removes exactly that scan, so the operation would end at r49
and those five ordinals would return `Ok` with a paragraph served. They are
*held* mutations: B2 refuses them and B1 would not. **The condition 0652 sets
for B1 is therefore not met, and this change stops at B2.**

## Breaking changes

No existing public signature changed. Two public items are added and one
documented contract moves.

| item | crate | kind | what it is |
| --- | --- | --- | --- |
| `SharedOleFile::caller_bracketed_identity_fingerprint(&self) -> Result<ArtifactFingerprint, OverlayError>` | `litchi-cfb` | **added** | the empty-splice identity digest in one complete scan; the composed CFB reopen is retained; the caller owns the read-twice-compare, and the rustdoc says so and says when the method may not be used |
| `SharedOleFile::unbracketed_source_fingerprint(&self) -> Result<ArtifactFingerprint, OverlayError>` | `litchi-cfb` | **added** | the complete source digest with no composed reopen and no comparison; the rustdoc states that it is a failure-path diagnostic and never an identity capture |
| `litchi_doc::body_text::source::SourceSnapshot::{source_version, fingerprint, len, is_empty, limits}` | `litchi-doc` | **contract documented and moved** | unchanged signatures; their rustdoc now states that they return state retained at open, that Option C ends the open's fence one identity call earlier, and that every path which yields a byte still goes through the same complete-artifact comparison |
| `litchi_doc::body_text::source::SourceSnapshot::open*` | `litchi-doc` | **behaviour** | six read ordinals of an open that were refused now return `Ok` and are refused by the first access instead (Option C); on the FAT-sector witness, one further ordinal changes its typed error (cost B-iii, below) |

No limit was relaxed, no `unsafe` added, no dependency added, no threading, no
ambient I/O. ADR 0001's layering, ADR 0003's authorization rule and ADR 0005's
no-leakage rules are untouched; no ADR text is amended, because 0652 assigns
this change none.

## Why it is sound

**The reduced call's comparison partner is a scan the operation already takes.**
`caller_bracketed_identity_fingerprint` returns a digest read once. What makes
it sound is not the entry point but where it is called:

* the generic open's first call is compared against `final_fingerprint`, the
  planning scan of identity call 2, by the `if` at `source.rs:611`;
* the owned open's first call has the same partner, and additionally cannot be
  mutated at all — `OwnedSource` is constructed inside the module and never
  handed out;
* `ensure_current`'s first call is compared **twice**: against the digest
  retained at open, and against `confirmed`, the second call's pair.

A caller with no such partner must keep using
`plan_same_length_stream_splices`, and the entry point's rustdoc says that in
those words. The PPT open is exactly such a caller, and it is why PPT is
untouched.

**The helper is the empty-splice case of the existing planner, not a parallel
path.** Its prelude is the one `plan_same_length_stream_splices` runs —
`check_source_version`, then `ensure_length`, then `fingerprints` — in the same
order, with the same `From` conversions. For an empty splice list the two things
it does not reproduce are unreachable: the per-stream comparison buffer
`plan_same_length_stream_splices` reserves, which no selection would read, and
its `OverlayLimits::new`, whose only failure is a zero limit that
`StreamSpliceLimits`' own constructor already rejects and whose fields are
private. Both are named here rather than left for a reader to rediscover.

**The composed reopen stays.** ADR 0003's proof that the candidate parses is
retained by both the reduced call and the paired one; Option D remains rejected.
`caller_bracketed_identity_still_reopens_the_composed_candidate` is the unit
test, and 14 of the FAT sweep's 17 `Overlay(Ole(..))` rows keep both wrapper and
site, which is the corpus-level test.

**Option C relocates a refusal; it does not lose one.** R5 and R6 bracketed one
`ensure_source_identity` and no consumed byte. The `ensure_source_identity` is
kept. Every ordinal they used to refuse is now refused by `ensure_current` at
the first access — measured, not modelled, in the sweep bands below.

**Cost B-iii is repaid, and the repair is narrow.** Removing the open's first
confirming scan hands the window in which a caller's writer corrupts the
artifact's own allocation table to the retained index parse, which would report
`CorruptedFile` and blame the file. `relocate_structural_cfb_error` runs only on
that parse's error, only for the four `OleError` variants that describe the
artifact's structure (`InvalidFormat`, `InvalidData`, `NotOleFile`,
`CorruptedFile` — a limit refusal, an I/O failure, an allocation failure and an
already-typed `SourceChanged` keep their own error exactly), and only on
positive proof: the digest is recomputed reopen-free and must **differ** from
the one the opening scan captured. If the recomputation itself fails, or the
artifact is unchanged, the original structural error stands. It cannot fire on
the composed reopen's own error, which is the inversion 0644's gate G1 fails an
implementation for.

**Why the re-hash needs its own entry point.** On the branch it runs from, the
allocation table has just been corrupted. `plan_same_length_stream_splices` and
`caller_bracketed_identity_fingerprint` would both reopen the composed candidate
over that damage and fail with a third, unrelated error instead of reporting
whether the bytes moved.
`only_the_unbracketed_digest_survives_an_index_damaged_under_the_read` is the
unit test of exactly that: the two bracketing entry points fail with
`Overlay(Ole(..))` and the digest-only one returns the moved digest.

**What is given up, stated precisely.** A *self-erasing* transient — a mutation
that appears after a reduced call's planning scan and is fully reverted before
the operation's next scan — is no longer observed inside the window the removed
scan covered. Every *held* mutation from the first complete scan onward is still
refused; the sweeps below assert that at every ordinal, on three witness
placements, over the whole corpus.

## Measured

Host AMD EPYC 9R45, 32 cores, Linux 7.0.0-1012-aws, rustc 1.95.0, `--release`,
every measured process `taskset -c 14`; eight agents were building concurrently,
so quiescence is **not** established and the A/A legs below are the in-window
floor. Before leg `/home/zhuhe/code/litchi-worktrees/before-70d7768cc`, after leg
this branch. The probe is change 0644's, reused verbatim (`probe/main.rs`
sha256 `35251d10…`); the publication differential is change 0589's, likewise.

### Deterministic counts (0644's gate G3)

Complete scans and read calls, traced on all 38 fixtures
([`trace/scans-before.tsv`](results/change-0659/trace/scans-before.tsv),
[`…-after.tsv`](results/change-0659/trace/scans-after.tsv)):

| operation | fixtures | complete scans | read calls |
| --- | ---: | --- | --- |
| DOC `open` | 8 | **6 → 3** on every one | 30 → 23 (6), 31 → 24, 46 → 34 |
| DOC `open` + `paragraph(0)`, admitted read | 4 | **14 → 9** | 59 → 50 (3), 60 → 51 |
| DOC `open` + `paragraph(0)`, refused read | 4 | **10 → 6** | 43 → 35 (3), 68 → 54 |
| PPT `open` | 30 | **2 → 2** | identical on every one |

Those are exactly the counts 0644's G3 requires for the declared variant B2,
and the PPT row is the test that Option B did not leak into the PPT open. Index
parses per DOC open fall 5 → 4 (I3c goes with identity call 3) and per PPT open
stay at 2, visible in the traced read order
([`trace/trace-documentProperties-doc-open-after.txt`](results/change-0659/trace/trace-documentProperties-doc-open-after.txt)).

**`version()` and `len()` are reconciled, not asserted unchanged.** On
`documentProperties.doc` a DOC open takes 108 `version()` and 32 `len()` calls
before and **73 / 22** after; on `picture.doc`, 132 / 32 before and **87 / 22**
after ([`counts/census-doc-delta.tsv`](results/change-0659/counts/census-doc-delta.tsv)).
Both reductions are exactly attributable to the removed work, with nothing left
over. One `fingerprints` call over an *n*-chunk artifact costs `4 + 2n`
`version()` and 2 `len()`; identity call 3 additionally costs its
`check_source_version` (1), its `ensure_length` (2 + 1) and its composed reopen
(`2 × reads + 6` `version()`, 3 `len()`), and I3c is 4 reads on
`documentProperties.doc` and 6 on `picture.doc`:

| fixture | removed | predicted `version` | measured | predicted `len` | measured |
| --- | --- | ---: | ---: | ---: | ---: |
| `documentProperties.doc` | 3 `fingerprints` + all of call 3 | 6 + (1+2+6+6+14) = **35** | 35 | 2 + (1+2+2+3) = **10** | 10 |
| `picture.doc` | the same | 8 + (1+2+8+8+18) = **45** | 45 | 2 + (1+2+2+3) = **10** | 10 |

The five `ensure_source_identity` sites (10 `version()`, 5 `len()`) are
untouched, as G3 requires, and the PPT counts do not move at all.

**Bytes read**, over the whole census
([`counts/census-doc-delta.tsv`](results/change-0659/counts/census-doc-delta.tsv),
[`census-ppt-delta.tsv`](results/change-0659/counts/census-ppt-delta.tsv)):

| corpus | before | after | reduction |
| --- | --- | --- | --- |
| 8 admitted `.doc`, `open` | 6.05×–7.47× the artifact | 3.04×–4.21× | median **−44.1%** (−43.7%…−49.8%) |
| 49 refused `.doc`, `open` | 0.06×–3.15× | 0.06×–2.15× | median **−42.3%** (0%…−49.3%) |
| 4 admitted `.doc`, `open` + `paragraph(0)` | 14.40×–16.64× | 9.36×–11.37× | median −31.8% |
| 30 `.ppt`, `open` | 2.02×–2.53× | **identical** | 0.0% |

The refused row is the saving 0644 said the −19.9% open median would not show,
because 49 of the 57 `.doc` fixtures are refused *inside* the semantic parse and
never reach identity call 2 or 3: Option C does nothing for them, and B2 is the
whole of their −42.3%. 0644 predicted 31%–49% there.

### Native cycles (0644's gate G4)

`perf stat -r 3 -e cycles,instructions` isolation pairs, legs in the order
A1 B1 B2 A2, each value the mean of its two legs
([`perf/summary-open.txt`](results/change-0659/perf/summary-open.txt)):

| fixture | bytes | before | after | Δ | 0644's predicted after | residual |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `documentProperties.doc` | 9,728 | 333,288 | 216,956 | **−34.90%** | 270,153 | −19.69% |
| `endingnote.doc` | 9,728 | 333,896 | 216,796 | −35.07% | 270,762 | −19.93% |
| `footnote.doc` | 9,728 | 333,954 | 217,382 | −34.91% | 270,820 | −19.73% |
| `noheadfoot-litchi.doc` | 10,240 | 340,895 | 220,344 | −35.36% | 274,438 | −19.71% |
| `lists-margins.doc` | 10,752 | 348,928 | 225,202 | −35.46% | 279,148 | −19.33% |
| `table-merged-cells.doc` | 17,408 | 429,232 | 263,490 | −38.61% | 316,256 | −16.68% |
| `duplicate-style-names.doc` | 64,512 | 1,038,393 | 571,184 | −44.99% | 619,714 | −7.83% |
| `picture.doc` | 1,448,448 | 18,974,070 | 9,590,476 | **−49.45%** | 9,573,734 | +0.17% |

**DOC open: median −35.41%, range −34.90% to −49.45%.** DOC `open` plus
`paragraph(0)` on the four fixtures whose paragraph read is admitted: median
**−24.96%**, range −24.59% to −32.23%
([`perf/summary-read.txt`](results/change-0659/perf/summary-read.txt)). **PPT
open, all 30 fixtures: median −0.12%, range −0.89% to +0.16%** — not one fixture
adverse by more than 0.16%, against G4's 5% bar. **A/A floor in the same window,
same counters: median +0.06%, range −0.84% to +0.76%, p95 of the absolute
deltas 0.62%** over 76 same-leg pairs.

**G4's ±5% band against 0644's predicted after-value fails on 7 of the 8 DOC
fixtures, and the reason is a term the prediction omitted, not a scan this
change removed.** 0644 modelled the saving as `k × 2.16331 × bytes` — the
per-byte hashing term alone. Fitting the *measured* saving as
`3 × c × bytes + fixed` by OLS over the eight fixtures
([`perf/fit-open.txt`](results/change-0659/perf/fit-open.txt)) gives

    saving = 6.4412 × bytes + 53,780 cycles      (fit residual −0.67%…+0.45%, median −0.06%)

so `c = 2.1471` cycles per byte per scan, which agrees with 0644's own OLS
constant of 2.1633 to **0.75%** across a 149× size range — the per-byte term is
exactly what 0644 measured, and exactly three scans were removed, which G3
confirms by counting them. The unmodelled **53,780 cycles** is the fixed term:
the composed CFB reopen I3c and the metadata observations that went with it. The
readback fit separates the two contributions, because it removes five
`fingerprints` calls and the same one reopen
([`perf/fit-read.txt`](results/change-0659/perf/fit-read.txt)):

    saving = 10.6317 × bytes + 72,996 cycles     (c = 2.1263; residual −1.53%…+1.34%)

Solving `3f + r = 53,780` and `5f + r = 72,996` gives **f ≈ 9,608 cycles** for
one removed `fingerprints` call's metadata observations and **r ≈ 24,956
cycles** for the removed composed reopen — against the 21,393 cycles 0644
measured for one *standalone* `SharedOleFile::open` on this fixture
(`results/change-0644/perf/perfstat-cfb-index.tsv`), which the composed reopen
exceeds because every one of its reads is bracketed by two `version()`
observations on the real source.

So the gate's purpose — *"a DOC result that beats the prediction by more than 5%
… would mean a further scan was removed"* — is **satisfied**: no further scan
was removed, the count says so on all 38 fixtures, and the excess is one fixed
term that fits the whole corpus to ±0.7%. The gate's literal threshold is
**not** met, and this record says so rather than restating the prediction.

### Paired timing

0589's `bench` driver, order A1 B1 B2 A2, 120 samples per leg (240 per
scenario), pinned to CPU 14, nanoseconds
([`bench/bench-summary.txt`](results/change-0659/bench/bench-summary.txt)):

| scenario | before p50 | after p50 | after vs before | before vs after | before p95 | after p95 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| DOC open, `documentProperties.doc` | 72,890 | 47,860 | **−34.34%** | +52.30% | 79,021 | 52,540 |
| DOC open, `lists-margins.doc` | 76,260 | 49,765 | −34.74% | +53.24% | 82,361 | 54,770 |
| DOC open, `duplicate-style-names.doc` | 229,662 | 126,120 | −45.08% | +82.10% | 236,661 | 132,761 |
| DOC open, `picture.doc` | 4,256,411 | 2,150,101 | **−49.49%** | +97.96% | 4,270,591 | 2,159,230 |
| DOC open + paragraph, `documentProperties.doc` | 157,551 | 120,601 | −23.45% | +30.64% | 164,151 | 126,941 |
| DOC open + paragraph, `lists-margins.doc` | 163,610 | 124,746 | −23.75% | +31.16% | 170,371 | 131,380 |
| DOC open + paragraph, `duplicate-style-names.doc` | 526,103 | 356,326 | −32.27% | +47.65% | 529,373 | 363,212 |
| PPT open, `45543.ppt` | 378,342 | 378,807 | **+0.12%** | −0.12% | 385,271 | 385,892 |
| PPT open, `SampleShow.ppt` | 134,250 | 134,531 | **+0.21%** | −0.21% | 140,801 | 141,231 |

**A/A floor in the same window: median +0.08%, range −0.70% to +0.84%, p95 of
the absolute deltas 0.71%** over 18 same-leg pairs — far under the 5% p50 bar,
so the timing is reported rather than set aside. Every scenario run is in the
table. The two PPT rows are the only ones that got *worse*, by +0.12% and
+0.21% at p50 — inside that floor, and beside a PPT cycles median of −0.12% and
read counts that are identical fixture by fixture.

## Correctness evidence

### Gate G1 — change-under-read sweeps

A held mutation is placed after **every** read ordinal of a clean run, on three
witness placements — outside every parsed range, inside the first FAT sector,
inside the first directory sector — over the 8 admitted `.doc` in two modes and
all 30 `.ppt`, on both legs: **144 sweeps per leg, 3,444 mutated opens on the before leg and
3,009 on the after**. The placements are derived from each artifact's own CFB header
([`scripts/cfboffsets.py`](results/change-0659/scripts/cfboffsets.py)); on
`documentProperties.doc` and `45543.ppt` they reproduce 0644's hand-chosen 520,
8,800 and the FAT sector of 381,028 exactly. The before leg reproduces every one of 0644's retained
outputs **byte for byte** at this base commit: its four traces, its six named
sweeps, its four transient witnesses, its `FileSource` witness and its
57-fixture census.

**The unconditional half passes.** Per sweep pair
([`sweeps/gate1-sweep-table.txt`](results/change-0659/sweeps/gate1-sweep-table.txt)):
the leading `OK` window (a mutation preceding the first complete scan, which
binds the snapshot to the bytes that now exist) is identical on both legs
everywhere; **no ordinal anywhere became `OK` in the interior**; and the
trailing `OK` window is **identical on both legs of all 52 DOC sweeps** — 1
ordinal on 39 of them, 2 on `picture.doc`'s FAT witness (a complete scan of that
1.45 MB artifact is two chunks and the witness byte lies in the first, so the
last read cannot see it, before or after), and 0 on the twelve `open` plus
`paragraph(0)` sweeps of the four fixtures whose paragraph read is refused for a
semantic reason, where every ordinal ends in that refusal. Cost B-ii's wider
window is B1's, and B1 was not taken.

**PPT is byte-identical.** All **92** retained PPT sweep files — 30 fixtures ×
3 placements plus 0644's two named ones — have the same sha256 on both legs
([`sweeps/manifest.tsv`](results/change-0659/sweeps/manifest.tsv)).

**The declared exceptions, and no others.** They were declared before the run
and are these two.

*Option C's relocation.* The open is six ordinals shorter, so the ordinals that
used to be refused by identity call 3 are now after the open's last read. On the
payload witness they are refused by the first access with the identical variant:
over `open` plus `paragraph(0)` the bands are `0–4 OK`, `5–29 open-ERR
SourceFingerprintChanged`, `30–58 read-ERR SourceFingerprintChanged`, `59 OK`
before and `0–4 OK`, `5–22 open-ERR …`, `23–49 read-ERR …`, `50 OK` after. On
the FAT witness they become `read-ERR Overlay(Ole(CorruptedFile))`, because the
corrupted sector makes the first `ensure_current`'s composed reopen fail before
a digest is compared — which is exactly what 0644's G1 says the FAT expectation
must be, and a gate demanding `SourceFingerprintChanged` there would be wrong
([`sweeps/bands-doc-corpus.txt`](results/change-0659/sweeps/bands-doc-corpus.txt)).

*Cost B-iii, now measured rather than derived.* 0644 predicted one affected
window of three ordinals. The FAT sweep over `documentProperties.doc` says the
truth is one deleted position and **one** re-typed ordinal:

| position of the held mutation | before | after |
| --- | --- | --- |
| after I1's third read | `Ole(CorruptedFile)`, 3 reads | unchanged |
| during I3a (2–6) | `Overlay(Ole(CorruptedFile))`, 8 reads | **unchanged**, 8 reads |
| after I3a, before the retained parse (7–9) | `Overlay(SourceFingerprintChanged)`, 10 reads | **same error, byte for byte**, 13 reads |
| after R2 (10) | `Ole(CorruptedFile)`, 13 reads | the position is R2; it no longer exists |
| after the retained parse's first read (11) | `Ole(CorruptedFile)`, 13 reads | **`Overlay(SourceFingerprintChanged)`** |
| during I3b (12–20) | `Overlay(Ole(CorruptedFile))`, 22 reads | **unchanged** (11–19), 21 reads |
| after I3b (21–23) | `Overlay(SourceFingerprintChanged)`, 24 reads | **unchanged** (20–22), 23 reads |
| identity call 3's own rows (24–29) | `Overlay(Ole)` ×3, `SourceFingerprintChanged` ×3 | relocated by Option C |

The one re-typed ordinal is the price of removing R2, and it is unavoidable
rather than a choice: R2 was the *only* observation that distinguished "the
mutation arrived before the retained parse" from "the mutation arrived after R2
confirmed the identity". With it gone the two are indistinguishable from inside
the program, and the relocation resolves both the same way — as a change of the
source, which is what happened. The class counts:
`Overlay(Ole(CorruptedFile))` **17 → 14**, which is precisely G1's requirement
that I3a's five rows and I3b's nine keep both wrapper and site;
`Overlay(SourceFingerprintChanged)` 9 → 7; `Ole(CorruptedFile)` 3 → 1.

**The relocated error is the same value, not merely the same variant.** The
sweep's rendering of rows 7–10 after is character-for-character the before
leg's row 7, including `expected` — because the relocation reports
`expected: opening_fingerprint`, which is the digest the removed confirming scan
compared against, and `observed`, which is a complete SHA-256 of the same
mutated bytes it would have hashed.

**The refusal counts are asserted exactly**, per fixture and per placement, in
[`sweeps/gate1-sweep-table.txt`](results/change-0659/sweeps/gate1-sweep-table.txt);
a lost fence point would shrink one rather than pass silently.

### Gate G2 — byte-identical publication

Change 0589's 87-artifact differential, rerun unchanged on both legs: for every
`.doc` and `.ppt` under `test-data/`, the empty-splice identity plan's two
fingerprints and `is_noop`, the SHA-256 of the bytes it publishes, the publish
report's two fingerprints, an exact byte no-op splice, an effective one-byte
splice with its span count and published digest, and the DOC and PPT snapshot
fingerprint or the exact typed refusal. **The two reports are byte-identical**
([`differential/`](results/change-0659/differential/digest-after.tsv)), and the
before leg reproduces 0589's own retained file byte for byte.

### The four transient witnesses

0644's W-T1…W-T4 were rerun on both legs, and the boundary was then mapped
rather than sampled, because the reduced call's partner moved
([`witness/`](results/change-0659/witness/)):

* **On the open's first identity call** — where B2 takes the scan — a flip after
  R1 reverted after read *k* is accepted for `k ≤ 9` before and for `k ≤ 17`
  after, and refused from `k = 10` before and `k = 18` after. The window in
  which a self-erasing transient goes unremarked grows from **4 reads to 12**:
  from "reverted before R2" to "reverted before R3". W-T1 (flip 5, revert 10),
  refused before, is accepted after; W-T2 (revert 9) is accepted on both legs.
  This is cost B-i, measured on the call that pays it.
* **Inside `ensure_current`** — where 0644's adopted B3 would have taken it —
  the window does **not** grow. A flip after that call's planning scan reverted
  after its last read is accepted on both legs (before r32–r35, after r25–r28),
  and reverting one read later is refused on both. The removed confirming scan
  was immediately followed by the next call's planning scan, so the partner sits
  at the same relative ordinal. **B costs nothing here**, which 0644 could not
  know because it did not run the changed code and said so under its own
  Limitations.
* The **`FileSource` witness** (W-A) is byte-identical on both legs and to
  0644's: an in-place `pwrite` through a second descriptor with the modification
  time restored leaves `SourceVersion` and `len()` unchanged. Option A stays
  rejected; nothing here substitutes a metadata observation for a scan.

### Tests added

`litchi-cfb` (`overlay_tests.rs`): the caller-bracketed digest equals the
empty-splice plan's `source_fingerprint()` and takes one complete scan where the
plan takes two; it still reports the composed reopen's structural failure; the
unbracketed digest is the same value with exactly one read and no reopen; and,
over an index damaged under the read, both bracketing entry points fail with
`Overlay(Ole(..))` while only the digest-only path reports the movement.

`litchi-doc` (`body_text::source::tests`): the generic open takes exactly 3
complete scans and `open` plus `paragraph(0)` exactly 9, with a second paragraph
read taking 6 more; a table mutated after the opening scan is still an identity
change carrying the clean fingerprint as `expected`; the relocation fires on
neither a file that was malformed before anything was read nor the composed
reopen's own error; and Option C's trailing ordinal is accepted by `open`,
reports the open-time values from the retained accessors, and is refused by both
`paragraph` and `edit_paragraph`. Change 0589's two existing per-ordinal sweep
tests pass unchanged and now exercise the relocated refusals.

### Gates

Run in `/home/zhuhe/code/litchi-worktrees/0659`; tails in
[`gates.txt`](results/change-0659/gates.txt).

| gate | result |
| --- | --- |
| `cargo fmt --all --check` | pass |
| `cargo clippy -p litchi-cfb -p litchi-doc -p litchi-ppt --all-targets` | pass, no warning |
| `cargo test -p litchi-cfb -p litchi-doc -p litchi-ppt` | 2,776 passed, 0 failed, 25 ignored |
| `cargo doc -p litchi-cfb -p litchi-doc -p litchi-ppt --no-deps` | pass |
| consumer crates of the three (`litchi-ole-common`, `litchi-vba`, `litchi-xls`, `litchi-crypto`, `litchi-sign`, `litchi-ograph`, `litchi-docx`, `litchi-pptx`) | 4,081 passed, 0 failed, 34 ignored |
| `cargo test -p litchi --features docx,xlsx,pptx,xls` | 265 passed, 0 failed, 7 ignored |
| `cargo test` in `tools/perf-baseline` | 540 passed, 0 failed, 1 ignored — the two allocator tests change 0651 lists as flaky did not fail |
| `python3 tools/non_iwork_gate.py verify` | pass (45 bulk tree roots, 35 facade-safe trees, 1 combined tree) |
| G1, G2, G3, G4 | above |

## Validation preserved

Every `ensure_source_identity`, `check_source_version` and `ensure_length` is
where it was; the open still has five of the first and `ensure_current` still
has three. No limit was relaxed and no malformed-input defence weakened: the one
fixture the corpus refuses before any complete scan, `redline-1.doc`, still
refuses after 512 bytes with `Ole(InvalidFormat("Invalid byte order"))`, and the
admitted and refused sets of the 57-fixture DOC census, the 57-fixture readback
census and the 30-fixture PPT census are identical on both legs in every
outcome. The complete-artifact SHA-256 remains the only thing that establishes
that two observations saw the same bytes; Options A, D, E, F and G stay
rejected. An exact byte no-op still publishes the source unchanged, and G2
proves no digest and no published byte moved.

## Limitations — what is not claimed

* **No speedup is registered as a claim.** `performance_claim: none`. The
  figures are scoped to this host, this corpus, warm cache, single thread,
  CPU 14, with eight agents building concurrently and the floors stated beside
  them.
* **G4's literal ±5% band is not met on the DOC open**, and the section above
  says so plainly. The evidence that no extra scan was removed is the count, not
  the cycles.
* **B1 is rejected on a witness drawn from the B2 sweeps**, not from a build of
  B1. Ordinals 45–49 of the landed `open` plus `paragraph(0)` are refused today
  and would fall outside B1's last scan; that is a held mutation B2 catches, so
  0652's condition fails. Nobody measured what B1 would cost or save here.
* **The relocation's proof that the bytes moved is a complete re-hash on a
  failure path.** It costs one more complete scan on exactly the inputs that are
  failing structurally *and* have changed under the read — 4 ordinals of one
  sweep on this corpus, none of them a benign input. It is not on the common
  path, and standing trade-off 3 of 0652 is what permits placing it there.
* **0644 says the owned DOC path is unaffected because
  `source_is_owned_immutable` already elides its confirming scan. That is
  incorrect at this base commit**, and this record acts on the code rather than
  the claim: `open_owned_source` builds its index through
  `SharedOleFile::open_with_limits`, which sets the flag to `false`, so the
  owned path's identity calls did take two scans each. It now takes B2's rule
  uniformly (6 → 3 scans), which is sound there twice over — the outer
  comparison exists *and* `OwnedSource` cannot be mutated by a caller. The owned
  path was not separately timed.
* **Cost B-iii is one re-typed ordinal on one fixture's FAT witness**, measured
  on `documentProperties.doc`. The corpus sweeps show the same band structure on
  the other seven `.doc`, but the per-ordinal attribution to a named call site
  is read from the traced read order, as 0644's was.
* **The transient boundaries were mapped on `documentProperties.doc` only.**
  The classes they name are structural; no boundary sweep was run on the other
  seven `.doc`.
* **No cold-cache, physical-device, range-source, remote, cross-platform,
  allocation or RSS result**, and no `.doc` above 1.45 MB or `.ppt` above
  1.34 MB. The per-byte term means the saving grows with size, so the corpus
  understates it; the fits are not extrapolated beyond the measured range.
* **The `FileSource` witness is Unix only.** The Windows version policy tracks
  different metadata and would need its own witness.
* **`litchi-ppt` was not edited**, so "PPT unchanged" is a measurement of the
  shared primitive's effect, not of a PPT change. The PPT readback path
  (`read_text`) was not swept; 0644's observation that PPT's `ensure_current`
  checks only version and length is still read from the source.

## Retained evidence

[`results/change-0659/README.md`](results/change-0659/README.md) — the contents
table, provenance with commit hashes and binary sha256s, every raw output cited
above, and the scripts that produced them.
