# 0644: four of the DOC snapshot's six complete reads are load-bearing after all — a `SourceVersion` can replace none of them, and removing R2 as well would cost a second public entry point to keep one error from changing its name

Status: **retained, design only. Nothing under `crates/` changed.**
`performance_claim: none` — the read inventory, the change-under-read sweeps,
the mutation witnesses, the native cycles and the per-scan fit below are
reported as evidence, not registered as a claim. The predicted savings are
modelled from a measured constant and are labelled as such.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This is queue item 10 of change [0630](0630-queue-refresh-after-the-first-wave.md). It
answers the question change [0589](0589-ole2-snapshot-fingerprint-passes.md)
left open — after 0589 halved the *hashing*, a generic-`ReadAt` DOC
`SourceSnapshot::open` still makes **six complete artifact reads** for three
identity passes — and the question change
[0609](0609-facade-doc-source-route-design.md) left open: whether the
source-backed route's `12.90 cycles per file byte` can be cut without weakening
ADR 0006's guarantee that a splice is applied only to the bytes it was planned
against.

The answer has four parts, and it is more conservative than the question
assumed. **Two of the six come out of the DOC open** — the third identity pass,
which brackets no consumed byte. **A third is reachable but costs a second
public entry point**, because the confirming scan of the open's first identity
call is what makes `SourceFingerprintChanged` outrank a CFB structural error at
one window of the corpus, and repairing that by change 0621's relocation rule
needs a reopen-free re-hash that the identity entry point cannot provide. **None
of the six can be replaced by a `SourceVersion` observation**, and the witness
that proves it runs on an ordinary `FileSource` in this repository. **The
reduction is not admissible on the PPT open** at all, because that open has no
outer bracket to inherit; the sweep shows which mutation it would stop
detecting. Where the reduction *is* free is the **readback**: `resolve`'s eight
complete scans become six, because there `ensure_current`'s two calls already
bracket each other and no index parse follows a reduced one.

No routing, no crate edit and no ADR amendment is proposed. The record freezes
the design, the witnesses and the admission gates so that an implementing change
starts from them.

**Against the queue item's own framing.** 0630 item 10 scopes the opportunity as
"the third identity pass and the duplicate index parse, **2 of the 6** surviving
complete reads". The adjudication is split, and the record says so rather than
reporting a larger number as if it were the item. The third identity pass is
**admitted** (Option C) and is 2 of the 6. The duplicate index parse is
**rejected** (Option E): it is not a complete read at all — it is 21,118–51,850
cycles across the 8 admitted DOC fixtures, 0.27%–6.45% of the open — and
removing it would move the only pre-scan structural refusal behind a complete
hash. So item 10's two reads become **exactly two**, from one of its halves and
not the other. A **third** and **fourth** are reachable from somewhere 0630 did
not name — the confirming scan inside each identity call, which the caller's own
outer comparison already duplicates — and both are recorded as priced variants
rather than folded into the headline, because the third costs a second public
entry point and the fourth costs a wider trailing window. The place the
confirming scan comes out for free is the readback, which item 10 did not
mention at all.

## What was changed

Nothing. `git diff --name-only c7326f680 -- crates/` is empty on this branch.
The change adds this record and the evidence packet under
`results/change-0644/`, whose `probe/` directory carries the retained scratch
driver that produced every number here.

## The measured inventory: what each of the six reads proves

Every read on the path was enumerated by a recording `ReadAt` adapter rather
than by grep, in the style change [0621](0621-xls-open-fence-count.md)
established. A read is a *complete scan* when it is one chunk of a sequential
pass over `0..len` in `overlay::fingerprints`' 1 MiB chunks
(`crates/litchi-cfb/src/overlay.rs:1345`); every other read belongs to a CFB
index parse or to the DOC semantic parse. `documentProperties.doc` (9,728 B)
takes 30 reads on `SourceSnapshot::open`; `picture.doc` (1,448,448 B) takes 46,
and the structure is identical
([`trace/`](results/change-0644/trace/)):

| reads | site | what it is |
| --- | --- | --- |
| r1–r4 | **I1** | `SharedOleFile::open_with_limits` (`source.rs:515`) — the first index parse |
| **r5** | **R1** | identity call 1, planning scan (`overlay.rs:1033`) |
| r6–r9 | **I3a** | composed-CFB reopen inside identity call 1 (`overlay.rs:1041`) |
| **r10** | **R2** | identity call 1, confirming scan (`overlay.rs:1058`) |
| r11–r14 | **I2** | `SharedOleFile::open_with_limits` (`source.rs:531`) — the reopen that binds the retained index inside the bracket |
| r15–r18 | — | the DOC semantic parse: FIB, table prefix, CLX, `WordDocument` range |
| **r19** | **R3** | identity call 2, planning scan |
| r20–r23 | **I3b** | composed reopen inside identity call 2 |
| **r24** | **R4** | identity call 2, confirming scan |
| **r25** | **R5** | identity call 3, planning scan |
| r26–r29 | **I3c** | composed reopen inside identity call 3 |
| **r30** | **R6** | identity call 3, confirming scan |

So the six complete reads are **three pairs**, one pair per
`identity_fingerprint` call (`source.rs:2089`), and the DOC open pays **five**
CFB index parses: I1 and I2 on the real source, plus one composed reopen inside
each identity call. 0589's inventory already says the reopen runs "once per
identity pass" (`0589:67`); what the trace adds is the total and the placement —
its read-count paragraph accounts for "the two index parses" and does not carry
the other three (`0589:198`).

The comparison graph is `R1 == R2`, `R3 == R4`, `R5 == R6` (each internal to one
call) and `R1 == R3`, `R3 == R5` (the DOC open's own checks, the `if`s at
`source.rs:593` and `:601`). **Between R4 and R5 no byte is consumed at all**: the trace shows
r24 and r25 adjacent, separated by 11 `version()` calls and 4 `len()` calls and
nothing else — the tail of identity call 2, `ensure_source_identity` at
`source.rs:600`, and the head of identity call 3. That is 0589's "P7 brackets no
work", now measured rather than read off the source.

| read | what it and no earlier read can detect | can a `SourceVersion` replace it | disposition |
| --- | --- | --- | --- |
| **R1** | nothing — it *is* the identity the snapshot is bound to | no (Option A) | **keep** |
| **R2** | a mutation that appears after R1, survives the composed reopen I3a, and is reverted before R3 | no | **kept under B3**; removed by B2 and B1, whose cost is witness W-T1 below and cost B-iii |
| **R3** | a mutation spanning I2 and the whole semantic parse — the load-bearing read | no | **keep** |
| **R4** | a mutation that appears after R3, survives I3b, and is reverted before R5 | no | **kept under B3 and B2** (it closes the open's last surviving call); removed by B1 |
| **R5** | a mutation spanning `ensure_source_identity`, which consumes no byte | no | **removed by Option C** — relocated to the first access |
| **R6** | a mutation that appears after R5, survives I3c, and is reverted before the open returns | no | **removed by Option C**, which removes its whole identity call |

The PPT open is a different shape and the trace says so: 8 reads, of which
r1–r3 are the publisher's index parse, **r4** is the only planning scan, r5–r7
are the composed reopen and **r8** is the confirming scan. It has one identity
call, no semantic parse, and — decisively — **no second identity call to
bracket the first**.

## The options

Each option states what it removes, the mutation it stops detecting with a
constructed witness, and the typed error that moves. Every witness in this
section was run; the raw outputs are in
[`sweeps/`](results/change-0644/sweeps/) and
[`witness/`](results/change-0644/witness/).

### Option A — replace a complete read with a `SourceVersion` observation. **Rejected.**

On a `FileSource`, `version()` is a metadata observation: the Unix policy hashes
(device, inode, length, mtime seconds, mtime nanoseconds)
(`crates/litchi-core/src/source/file.rs:345-358`). The type's own rustdoc says
what that costs: *"Neither policy hashes file contents … a writer capable of
restoring all tracked metadata can evade detection … Callers that do not trust
other writers must additionally validate the bytes they consume."*

**Witness W-A** ([`witness/filesource-version-witness.txt`](results/change-0644/witness/filesource-version-witness.txt)),
run on a copy of `documentProperties.doc`: open a `FileSource`, record
`version()` and `len()`, then `pwrite` one flipped byte at offset 0 through a
second descriptor on the same inode, `fsync`, and restore the modification and
access times with `File::set_times`. The result:

```
byte_before     208      byte_after      47       bytes_changed    true
length_before   9728     length_after    9728     length_changed   false
version_before  SourceVersion { id: 1, revision: 0 }
version_after   SourceVersion { id: 1, revision: 0 }
version_changed false
```

Every byte of the artifact may be rewritten this way. The typed error that would
move is **every** `SourceFingerprintChanged` on the path: it would simply not be
raised, and the snapshot would serve a paragraph resolved through a layout
derived from bytes that no longer exist. That is the exact failure ADR 0006's
source-identity rule and ADR 0005's *"mutation during a read returns
`SourceChanged`"* exist to prevent.

This is also already settled policy rather than a fresh judgement. Change 0175
withheld the elision from generic sources for this reason, and change 0587's
survey classifies the substitution as *"an ADR 0006 policy amendment, not an
optimization"* (`0587-remaining-opportunity-survey.md:988-990`). A performance
record may not make it. **Answer: no, for all six reads, on a `FileSource` as
much as on a hostile adapter.**

The cheap check is nonetheless **retained everywhere it already is**.
`ensure_source_identity` (`source.rs:2061`) runs at all five of the DOC open's
fence points (`source.rs:519`, `:527`, `:535`, `:592`, `:600`) and three more
inside `ensure_current`, and it is necessary — it is what turns a length change into
`OverlayError::Unavailable` and a token move into `OverlayError::SourceChanged`
before any digest is compared. It is never sufficient, and no option below makes
it carry weight it cannot carry.

### Option B — one complete scan per empty-splice identity call. **Admitted as B3, inside DOC `ensure_current` only; B2 and B1 priced and not recommended; refused on the PPT open.**

`identity_fingerprint` reaches `finish_overlay_plan_with_owner`, which scans the
artifact, reopens the composed candidate, and — because a `FileSource` is not
`source_is_owned_immutable` — scans it again and compares
(`overlay.rs:1022-1070`). Option B adds an empty-splice-only entry point in
`litchi-cfb` that performs the planning scan and the composed reopen and
**returns the digest without the confirming scan**, leaving the read-twice-compare
to the caller's own outer bracket. It is admissible exactly where such a bracket
exists.

*On the DOC open it exists*: R1's partner is R3 and R3's is R5, both compared by
`finish_open` itself. *On DOC `ensure_current`* (`source.rs:764`) it exists:
`observed` and `confirmed` are two calls that already compare with each other,
so the pair collapses from four complete reads to two. *On the PPT open it does
not exist*: `text_edit::SourceSnapshot::open` makes one identity call and
nothing else (`text_edit.rs:422-429`).

**What B stops detecting, and the witness.** R2's unique detection is a mutation
that appears after R1 and is reverted before R3.

**W-T1** ([`witness/transient-doc-flip5-revert10.txt`](results/change-0644/witness/transient-doc-flip5-revert10.txt)) —
flip byte 8703 after read 5 (R1), revert it after read 10 (R2), stable version
token throughout:

```
reads 10   mutations_applied 2   restored_byte_identical true
outcome    open-ERR Overlay(SourceFingerprintChanged { … })
```

Today that is refused. Under Option B it would be accepted, and the accepted
snapshot would be bound to the bytes that exist when the open returns, because
the source is byte-identical to its original by then. **The typed error that
moves is `OverlayError::SourceFingerprintChanged`, and it moves to not being
raised at all — for a mutation with no persistent effect.**

**W-T2** ([`witness/transient-doc-flip5-revert9.txt`](results/change-0644/witness/transient-doc-flip5-revert9.txt))
is the control that bounds the claim. Flip after read 5, revert after read 9 —
one read earlier, so the mutation spans the composed reopen but not R2:

```
reads 30   mutations_applied 2   restored_byte_identical true
outcome    OK
```

The current fence *already* accepts a transient in the immediately neighbouring
window. R2 does not detect "a transient during the reopen"; it detects "a
transient that happens to still be present one read later". Option B does not
open a class of undetected mutation, it widens by one read the window in which a
self-erasing transient goes unremarked — and for the source kind that matters
this is already the documented envelope: `FileVersionPolicy`'s rustdoc says that
"a transition fully reverted between observations is not visible"
(`crates/litchi-core/src/source/file.rs:17-19`).

**What B does not touch.** Every *persisting* mutation stays detected. The
change-under-read sweep
([`sweeps/sweep-documentProperties-doc-open.tsv`](results/change-0644/sweeps/sweep-documentProperties-doc-open.tsv))
places a held mutation after each of the 30 read ordinals in turn:

| trigger | reads consumed | today's outcome | sole detecting comparison | under **B3+C** | under B2+C | under B1+C |
| --- | ---: | --- | --- | --- | --- | --- |
| 0–4 | 30 | `OK` | none — the mutation precedes R1, so the snapshot binds to the bytes that now exist | unchanged | unchanged | unchanged |
| 5–9 | 10 | `SourceFingerprintChanged` | **R1 ≠ R2**, internal to call 1 | **unchanged** — B3 leaves the open alone | `R1 ≠ R3` instead | `R1 ≠ R3` instead |
| 10–18 | 24 | `SourceFingerprintChanged` | R1 ≠ R3, the `if` at `source.rs:593` | unchanged | unchanged | unchanged |
| 19–23 | 24 | `SourceFingerprintChanged` | **R3 ≠ R4**, internal to call 2 | **unchanged** | **unchanged** — the last call keeps its pair | `R1 ≠ R3` instead |
| 24 | 30 | `SourceFingerprintChanged` | R3 ≠ R5, the `if` at `source.rs:601` | removed by C, relocated to first access | same | same |
| 25–29 | 30 | `SourceFingerprintChanged` | **R5 ≠ R6**, internal to call 3 | removed by C, relocated to first access | same | same |
| 30 | 30 | `OK` | the trailing control: after the open's last read, nothing can bracket it | 1 ordinal, unchanged | 1 ordinal, unchanged | grows to 5 — cost B-ii |

The `reads consumed` column comes from the sweep file. The split of the 24-read
band into `R1 ≠ R3` and `R3 ≠ R4`, and of the 30-read band into `R3 ≠ R5` and
`R5 ≠ R6`, is derived from the traced read order, since both members of a band
report after the same read; the FAT-sector sweep of Option D below reports the
same split independently, at its own ordinals.

The three bolded rows are the three *internal* comparisons. **Under the adopted
design B3 the open keeps all of them that survive Option C**: only call 3 goes,
so this table's `doc-open` column changes in exactly the two rows Option C
relocates. B2 additionally removes call 1's `R1 ≠ R2`, and B1 call 2's
`R3 ≠ R4` as well. **For a held mutation every removed comparison is covered by
the next outer one**, because a held mutation persists into it: under B2 one in
5–9 is still present at R3, and under B1 one in 19–23 is still present at the
scan that replaces R3's pair. No variant therefore loses a *refusal* on the
payload witness. What B costs is three other things, and the record prices each
rather than absorbing it — and it is the third that keeps B out of the open in
the adopted design.

**Cost B-i — the self-erasing transient.** Whichever identity call is reduced,
the detection given up is a mutation that appears after that call's planning scan
and is reverted before the operation's next one. W-T1 and W-T2 above demonstrate
the class on the **open's** first call, which is where B2 and B1 would take it.
For the adopted design the call reduced is `ensure_current`'s first, and the
witness was run there too
([`witness/transient-docread-flip31-revert36.txt`](results/change-0644/witness/transient-docread-flip31-revert36.txt),
[`…-revert35.txt`](results/change-0644/witness/transient-docread-flip31-revert35.txt)):
on `open` plus `paragraph(0)`, flipping byte 8703 after **r31** — the first
`ensure_current`'s first planning scan — and reverting it after **r36**, that
call's confirming scan, is refused today with
`read-ERR Overlay(SourceFingerprintChanged)` at 36 reads, with
`restored_byte_identical true`. **W-T4** is the control: reverting one read
earlier, after r35, is **already accepted** today and the run completes in 59
reads. So on the readback exactly as on the open, the confirming scan detects not
"a transient during the reopen" but "a transient still present one read later",
and what B gives up is a mutation that leaves no trace.

**Cost B-ii — the trailing window, and why a second variant exists.** The
composed reopen is built from the target fingerprint (`overlay.rs:1035-1041`),
so it can only run *after* a scan, never before one. With the confirming scan
gone, an identity call therefore *ends* in a reopen. Under B applied to every
call, plus C, the DOC open is 22 reads ending `R3' (r18) → I3b (r19–r22)`, and a
payload mutation landing at any of r18–r22 is seen by no scan: the trailing
window grows from **1 ordinal to 5** on `documentProperties.doc`. Those
mutations are still refused at the first access, by the same `ensure_current`
the readback sweep measures — the same relocation Option C makes, not a new
class — but it is a larger relocation than C's alone and the record does not
hide it inside B.

The exception has to be **uniform across the operation, not special-cased to the
open**, because `resolve`'s trailing `ensure_current` has the same shape: its
last identity call would also end in a reopen, and `open` plus `paragraph(0)`
would end in four reads no scan follows — a 5-ordinal window there too. The rule
that holds the window at 1 everywhere is therefore *B applies to every
empty-splice identity call except the last one of its operation*.

That gives three variants, and the numbers for all three are under *Predicted
saving*:

| variant | identity calls that take one scan | open scans | `open`+`paragraph(0)` | trailing window | pays cost B-iii |
| --- | --- | ---: | ---: | --- | --- |
| **B1** | every call | 6 → **2** | 14 → **6** | grows to 5 ordinals | yes |
| **B2** | every call except the last of its operation | 6 → **3** | 14 → **9** | **1 ordinal, as today** | yes |
| **B3** | only inside `ensure_current`, never the open | 6 → **4** | 14 → **10** | **1 ordinal, as today** | **no** |

**None of the three serves a wrong value.** Under all of them the retained
`Layout` and digest describe bytes the surviving scans bracket, and the reads
after an operation's last scan belong to a composed reopen whose candidate is
dropped — `finish_overlay_plan_with_owner` returns `owner = None` for an empty
span list and `candidate` is a local (`overlay.rs:1041-1052`). The consumed
bytes of `resolve` are the clearest case: the paragraph text is read at r43–47,
*before* either scan of the trailing `ensure_current`, and `observed` is
compared against the digest retained at open (`source.rs:768`), so the bracket
over those bytes closes whatever happens afterwards. The choice between the
variants is therefore about **when a refusal is reported and how much machinery
it costs**, not about correctness.

**Cost B-iii — error precedence, and what repairing it actually costs.** This
is the cost a naive reading of "remove the second scan" misses, and Option D's
FAT-sector sweep is what exposes it. At trigger 7 today the sole detector is
`R1 ≠ R2`, reporting `Overlay(SourceFingerprintChanged)`. Delete R2 and the next
thing to read that FAT sector is **I2** (`source.rs:531`), which reports
`Ole(CorruptedFile("MiniFAT chain exceeds its declared length"))` — a different
`litchi_doc::Error` variant, and one that blames the file for damage the
caller's own writer did. Under B2 that is the one window affected, because call
2's pair survives and call 3 is removed by C, which relocates 27–29 rather than
rewriting them; under B1 rows 21–23 join it.

Change 0621's rule for exactly this case is that an observation which outranked
another failure is **relocated onto the error branch rather than deleted**.
Applied here the relocation is real but narrower and more expensive than it
first looks, and three things pin it down.

* **It must not fire on the composed reopen's own error.** At FAT-sweep triggers
  5 and 6 the mutation lands after R1, so R1 hashed clean bytes and I3a's third
  read finds the flipped sector: today those rows are
  `Overlay(Ole(CorruptedFile))` at 8 reads. A scan relocated onto *that* branch
  would find the digest moved and report `SourceFingerprintChanged`, inverting
  two of the rows gate G1 pins. (Triggers 2–4 are safe: the mutation precedes
  R1, so the digest has not moved.) **The relocation is scoped to structural CFB
  errors from the retained index parse — I2 — and to nothing else.**
* **It must not fire on a semantic refusal.** The 48 refused `.doc` fixtures
  fail in `reject_container_components` and `read_fib` with
  `Refused(Drawing)`, `Refused(Field)`, `Refused(Encrypted)` and the like, not
  with a CFB structural error. Firing there would put the second scan back on
  84% of the corpus and erase the refusal-path saving.
* **It needs a reopen-free re-hash, which is a second public addition.** The
  repair has to sit on I2's error branch in `litchi-doc`, because that is where
  the 7–9 window is reported and the `litchi-cfb` entry point has already
  returned `Ok`. `first_shared` is still in scope there, so the branch is
  reachable — but it cannot call the new entry point or `identity_fingerprint`,
  because either would run its own composed reopen over the same corrupted FAT
  sector and fail with a **third** variant, `Overlay(Ole(CorruptedFile))`,
  instead of the `SourceFingerprintChanged` the relocation exists to produce.
  What it needs is a digest-only path — a public wrapper over the private
  `fingerprints` (`overlay.rs:1345`) with no reopen. That is a **second** public
  item in `litchi-cfb`, and one whose whole purpose is to hand out an
  unbracketed, un-reopened digest, which is the shape this record warns about
  everywhere else. Its rustdoc would have to say that it is a failure-path
  diagnostic and never an identity capture.

**That third bullet is why B3 exists and why it is the recommendation.** B3
leaves the DOC open's identity calls alone and applies B only inside
`ensure_current`, where no retained index parse follows a reduced call and so
**cost B-iii does not arise at all**: no error-branch re-hash, no second public
item, no scoping rules to get right. It gives up one scan of the open against
B2 — a median −13.2% instead of −19.9% — and `docs/GOAL.md`'s "make the smallest
coherent change" and "revert speculative complexity" decide that trade. B2 is
recorded in full so that an implementing change which judges the second public
item worth 6.7 percentage points can take it with the three bullets above in
front of it, and gate G1 tests the relocation directly if it does.

**Why PPT is refused.** With one identity call and no bracket, a single scan
would leave the PPT open's digest uncompared. The PPT sweep
([`sweeps/sweep-45543-ppt-open.tsv`](results/change-0644/sweeps/sweep-45543-ppt-open.tsv))
shows triggers 4–7 refused today with `Source(SourceFingerprintChanged)` by the
internal R1 ≠ R2 comparison and nothing else. Under Option B those four ordinals
would open successfully, and PPT's `ensure_current` (`text_edit.rs:515-530`)
checks only the version token and the length — 0589 records the same asymmetry —
so the mutation would not be caught on the read path either. It would surface
only at commit (`text_edit.rs:915`, `:990`). That is not a relocated refusal; it
is a `read_text` that returns characters decoded from bytes the retained
identity does not cover. **Option B is therefore scoped to callers that supply
the outer bracket, and the entry point's rustdoc must say so.**

### Option C — drop the DOC open's third identity call. **Admitted, with a stated relocation.**

This is 0589's designed-and-not-implemented item, now with the ordinals
measured. R5 and R6 bracket `ensure_source_identity` and nothing else; the trace
shows r24 and r25 adjacent.

**What C stops detecting, and the witness.** Sweep triggers 24–29: a mutation
that first becomes visible after identity call 2 completes. Today
([`sweeps/sweep-documentProperties-doc-open.tsv`](results/change-0644/sweeps/sweep-documentProperties-doc-open.tsv))
those six ordinals are refused at `open`. Under C the open's last read is r24,
so they become "after the open's last read" and the open returns `Ok`.

They are **relocated, not lost**. The readback sweep
([`sweeps/sweep-documentProperties-doc-open-read.tsv`](results/change-0644/sweeps/sweep-documentProperties-doc-open-read.tsv))
measures the receiving fence: over `open` plus `paragraph(Position::new(0))` the
run takes 59 reads, and **every** trigger from 30 to 58 — that is, every ordinal
after the open's last read and before the readback's last — is refused with
`read-ERR Overlay(SourceFingerprintChanged)`, the identical variant, raised by
`ensure_current`'s two passes around `resolve`. Triggers 24–29 fall into the
same class under C, so the refusal survives with the same typed error at the
first call that consumes a byte. That step is **modelled**, anchored on the
measured behaviour of the immediately adjacent window.

**The contract move C does make** is narrower and more precise than "the moment
of detection shifts". `SourceSnapshot` exposes five accessors that return
retained state and consume nothing: `source_version`, `fingerprint`, `len`,
`is_empty`, `limits` (`source.rs:689-715`). Under C a caller that opens and then
reads `fingerprint()` receives a digest for bytes that may no longer exist,
where today the open would have refused. **Every path that yields a byte —
`paragraph`, `edit_paragraph`, `transaction`, `edit` — goes through `resolve`
and therefore `ensure_current`, so no byte is served past a lost fence.** An
implementing change should say this in the accessors' rustdoc rather than in a
record.

### Option D — drop the composed-CFB reopen for an empty splice list. **Rejected.**

The reopen (I3a/I3b/I3c) is ADR 0003's proof that the composed candidate parses.
With no span the candidate is the source, which I1 and I2 already parsed, so the
reopen looks redundant. Two witnesses say it is not.

**W-D1, PPT** ([`sweeps/sweep-45543-ppt-open-index-region.tsv`](results/change-0644/sweeps/sweep-45543-ppt-open-index-region.tsv)) —
the same sweep with the witness byte moved into the CFB **allocation-table**
region: offset 381,028 falls in FAT sector 743, the first of `45543.ppt`'s six
FAT sectors (bytes 380,928–383,999; the directory chain starts at byte 384,000).
Triggers **1–5** are refused with
`Source(Ole(CorruptedFile("regular stream chain exceeds its declared length")))`.
Trigger 1 is the publisher's own index parse failing at its third read; the
read counts show that for triggers **2–5** — after that parse has completed —
the composed reopen is the only detector. The PPT open
retains the index parsed at r1–r3, **before** its only scan, so that index is
outside every bracket; the composed reopen at r5–r7 is currently the only thing
that re-reads it. Removing it would let a snapshot retain an index describing
pre-mutation bytes while its digest covers post-mutation bytes.

**W-D2, DOC** ([`sweeps/sweep-documentProperties-doc-open-fat-region.tsv`](results/change-0644/sweeps/sweep-documentProperties-doc-open-fat-region.tsv)) —
witness byte 520, inside the FAT sector, so a flip corrupts the chain. The
sweep resolves which parse reports it, and the wrapper differs by parse:

| trigger | reads | outcome |
| --- | ---: | --- |
| 1 | 3 | `Ole(CorruptedFile("MiniFAT chain exceeds its declared length"))` — I1 |
| 2–6 | 8 | `Overlay(Ole(CorruptedFile(…)))` — **I3a** |
| 7–9 | 10 | `Overlay(SourceFingerprintChanged)` — R1 ≠ R2 |
| 10–11 | 13 | `Ole(CorruptedFile(…))` — I2 |
| 12–20 | 22 | `Overlay(Ole(CorruptedFile(…)))` — **I3b** |
| 21–23 | 24 | `Overlay(SourceFingerprintChanged)` — R3 ≠ R4 |
| 24–26 | 28 | `Overlay(Ole(CorruptedFile(…)))` — **I3c** |
| 27–29 | 30 | `Overlay(SourceFingerprintChanged)` — R5 ≠ R6 |

The composed reopens are the reporting site for **17 of the 31 ordinals** —
I3a's 5, I3b's 9 and I3c's 3 — and they report
`Error::Overlay(OverlayError::Ole(..))` while I1 and I2 report `Error::Ole(..)` — different variants of `litchi_doc::Error` carrying the same
CFB message. (As in the sweep table above, the `reads` column is measured and
the attribution of each band to a named site follows from the traced read
order.) Removing them would rewrite the wrapper for those ordinals, and that is an
error-identity change.

It is rejected because what it buys is the smaller half of the prize and the
only half that costs a contract. One standalone `SharedOleFile::open` is
**21,118–51,850 cycles** across the 8 admitted DOC fixtures
([`perf/perfstat-cfb-index.tsv`](results/change-0644/perf/perfstat-cfb-index.tsv));
taking that as the estimate for one composed reopen — the composed view adds a
`read_at` wrapper and a derived version token over the same reads, so this is an
estimate and not a measurement of the reopen itself — the two reopens that
survive the adopted design are **0.55%–12.89% of the open's measured cycles**,
against B3+C's **12.5%–33.0%** and B2+C's 18.8%–49.5%:

| fixture | open, measured | two reopens, estimated | D | B3+C |
| --- | ---: | ---: | ---: | ---: |
| `documentProperties.doc` | 332,564 | 42,786 | 12.87% | 12.7% |
| `duplicate-style-names.doc` | 1,034,836 | 45,150 | 4.36% | 27.0% |
| `picture.doc` | 19,006,126 | 103,700 | 0.55% | 33.0% |

The 12.89% upper bound is `endingnote.doc`'s, marginally above
`documentProperties.doc`'s 12.87%; the three rows above are the extremes plus
one middle fixture, and the full eight are in
[`perf/predicted-savings.txt`](results/change-0644/perf/predicted-savings.txt).

On the smallest fixture D and the adopted design are within 0.2 points of each
other, and that is the fairest statement of the trade: D's prize is comparable
only where the artifact is small, and it shrinks to a fortieth of the design's
on the largest.

**Rejected: it is the only option here that would rewrite a typed error with no
way to relocate the precedence at all — the reopen's `Overlay(Ole(..))` has no
later observation to hand it to, where B's confirming scan at least has a
reopen-free re-hash available at a price — and its prize shrinks with artifact
size exactly where the fence costs most, from 12.87% of a 9.7 KB open to 0.55%
of a 1.45 MB one.**

Option B keeps the reopen, so **14 of those 17 rows keep their reporting
site** — I3a's and I3b's. The other three, rows 24–26, are I3c's, and I3c dies
with identity call 3 in *every* variant, because that is Option C: those three
relocate to the first access along with rows 27–29, exactly as the payload
witness's 24–29 do. And the table is not otherwise untouched either: the
`SourceFingerprintChanged` rows 7–9, 21–23 and 27–29 are reported *by* the very
scans B removes, which is cost B-iii. Gate G1 pins the whole table, band by
band, per witness placement.

### Option E — hash before the first index parse, so one index parse serves both the bracket and the semantic parse. **Rejected.**

This is the option that would let the semantic parse's directory reads be
*accounted into* the fingerprint bracket rather than requiring a second index
parse. They can be, and the mechanism is ordering, not tooling: I2 exists only because I1's
bytes were read before the first scan and are therefore outside every bracket
(`source.rs:528-530`). Reorder to *scan, parse, scan* and the single index parse
falls inside `R1 == R3` by construction. I1 then has no fingerprint role.

It is rejected because I1 is also a malformed-input defence, and the measurement
prices what removing it would cost. The census of every `.doc` under
`test-data/` ([`counts/census-doc.tsv`](results/change-0644/counts/census-doc.tsv)),
57 fixtures, classified by how much of the artifact was read before the refusal:

| class | fixtures | bytes read before refusal |
| --- | ---: | --- |
| refused before the first complete scan | **1** | `redline-1.doc`, 512 of 8,319 bytes (0.06×), 1 read |
| refused after at least one complete scan | 48 | 2.0×–3.2× the artifact |
| admitted | 8 | 6.05×–7.47× the artifact |

Only one fixture is refused by I1 today, and the typed error is
`Ole(InvalidFormat("Invalid byte order"))` — a file whose CFB signature is
wrong. Under Option E that refusal would arrive after a complete SHA-256 of the
artifact. The corpus understates the cost because the corpus is small: the DOC
snapshot's default ceiling is `SharedOleFileLimits::MAX_INPUT_BYTES`
(`crates/litchi-cfb/src/shared.rs:28`, an alias of `OleFileLimits::MAX_INPUT_BYTES`
at `crates/litchi-cfb/src/file.rs:361`), which is **2 GiB**, so a 2 GiB file whose first eight bytes
are not a compound-file signature would cost about `2 GiB × 2.16 = 4.6 × 10⁹`
cycles before the same 512-byte refusal — an amplification of roughly seven
orders of magnitude, on input an attacker chooses. `docs/GOAL.md` forbids
weakened malformed-input defences, and this is one.

There is a second, sharper witness. For a source that is **both** malformed and
mutating, the typed error changes rather than merely arriving late: today I1
reports the structural error before any scan, while under Option E the scan's
own `ensure_length`/`ensure_current` bracket would report `SourceChanged` or
`Unavailable` first. **Rejected**, and the prize is the same term Option D was measured on: the one
index parse it would save is 21,118–51,850 cycles across the 8 admitted DOC
fixtures (20,244–51,850 over all 38), which is **0.27%–6.45%** of the DOC open.

Option B keeps I1 exactly where it is, in its full form rather than as a
header-only precheck, so that every structural refusal keeps its typed error and
its ordinal. That is a deliberate choice: a header-only gate would relocate the
FAT and directory refusals past the scan, which is Option E in miniature.

### Option F — tee the index parse off the fingerprint scan. **Rejected.**

Since `fingerprints` already brings every byte through a 1 MiB buffer, a single
forward pass could in principle hand the index builder the sectors it needs and
place them in the bracket for free. It cannot in general. The CFB index is a
pointer chase whose read order is data-dependent and not monotone in file
offset: the DIFAT's overflow sectors are discovered one from the next, the FAT
sectors sit wherever the DIFAT says, and the directory chain is only resolvable
once the FAT is complete. A forward pass would have to retain whatever it might
later need — unbounded in the worst case, which ADR 0005's bounded-resources
rule forbids — or refuse files whose sector layout is not forward-compatible,
which narrows admission. The trace also prices the prize: the index parses and
the semantic parse together are **0.05×–1.47×** of the artifact against the
scans' 6.00× (`counts/census-doc.tsv`, `full_artifact_reads` column), so on the
fixtures where the per-byte term hurts most the whole prize is under 1%.

### Option G — defer the identity to the first access. **Rejected.**

Change [0165](changes/0165-doc-lazy-fingerprint.md) made the *owned* editor's
FNV-1a fingerprint lazy, and the analogy invites deferring this one. It does not
transfer. 0165's value was diagnostic — *"the FNV value is not an authorization
boundary"* — and exact byte equality authorized every apply. Here the digest
**is** the boundary `ensure_current` compares against, and deferring it would
leave the semantic parse (I2, FIB, CLX, `parse_clx`) bracketed by nothing: the
first `ensure_current` would capture the identity from post-mutation bytes and
compare it against itself, so the retained `Layout` would describe one set of
bytes while the identity covered another. That is not a relocated refusal but a
wrong-value path, and no witness is needed to reject it.

## The design

**Adopt Option B3 and Option C. Reject A, D, E, F and G. Record B2 and B1 as
priced alternatives.** The resulting DOC open is

    I1 → R1 → I3a → R2 → I2 → semantic parse → R3 → I3b → R4 → compare(R1, R3)

— the open's own identity calls are untouched, Option C removes the third, and
the composed reopen stays where it is because Option D is rejected. That is
**four complete scans** (from six) and **four index parses** (from five: I1,
I3a, I2, I3b, with I3c gone with its identity call), the same five
`ensure_source_identity` points, the same trailing window of one ordinal, and
**the same `Error` variant at every ordinal that carries one** — because no
retained index parse follows a reduced call, cost B-iii never arises and no
error-branch machinery is needed.

That is **24 reads** on `documentProperties.doc` against 30 today, ending at
r24 — a compared scan, exactly as today's r30 is.

`ensure_current` is where Option B is applied, and the reason it is safe there is
structural: the only `SharedOleFile::open` calls in `litchi-doc`'s source path
are at `source.rs:515`, `:531` and `:649`, all inside `open_with_options` or the
owned path. **No index parse runs inside `ensure_current` or anywhere else in
`resolve`** — `resolve_paragraph` and `reject_selected_property_revisions` read
streams through the index already retained at open, under `check_source_version`
alone — so no retained parse follows a reduced call and cost B-iii cannot arise.
Its first identity call takes one scan and its second keeps its pair, so
`ensure_current` costs **three** complete scans instead of four (11 reads) and
still ends in a compared scan. `resolve` therefore falls from eight complete
scans to **six** (27 reads, ending in a scan), and `open` plus `paragraph(0)`
from fourteen to **ten** (51 reads against 59, ending in a scan). The PPT open is
unchanged in every respect.

**B2** additionally applies B to the DOC open's first identity call — three open
scans instead of four, a median −19.9% instead of −13.2% — at the price of cost
B-iii's repair: a second public item in `litchi-cfb` whose only job is an
unbracketed digest, plus an error-branch re-hash at `source.rs:531` scoped to
structural CFB errors alone. **B1** drops the last-call exception as well: two
open scans, a median −26.5%, and a trailing window that grows from 1 ordinal to
5 on both measured modes. Both are recorded and priced; an implementing change
that wants either owes its own record the argument.

Stated against 0589's own naming: **the pair (P6, P7) — four complete reads —
becomes two under the adopted design, and one under B1**, and in no variant does
it become "one read plus a length-and-version check". The length-and-version
check is already present at every one of those points and stays; Witness W-A
proves it can never be what makes a reduction sound. What makes P7's removal
sound is that R5 and R6 bracket no consumed byte; what makes B's removals sound,
wherever it is applied, is that the reduced call's comparison partner is a scan
the operation already takes.

## Predicted saving

Measured first: the native cost of one complete scan. `perf stat -r 3`
isolation pairs on CPU 19, differenced over 1,000 (or 100) operations, for the
8 admitted `.doc` fixtures and all 30 `.ppt`
([`perf/perfstat-legA.tsv`](results/change-0644/perf/perfstat-legA.tsv)). A DOC
open takes 6 complete scans and a PPT open 2, so each format yields the per-scan
constant independently:

| format | estimator | model | scans | cycles/byte/scan | residual, median and range |
| --- | --- | --- | ---: | ---: | --- |
| DOC, 8 fixtures | **OLS** | `205,152 + 12.9799 × bytes` | 6 | **2.1633** | −0.26%, −1.28%…+0.74% |
| DOC, 8 fixtures | two-point | `206,301 + 12.9793 × bytes` | 6 | 2.1632 | +0.00%, −0.94%…+0.85% |
| PPT, 30 fixtures | **OLS** | `68,254 + 4.3487 × bytes` | 2 | **2.1743** | −0.66%, −4.06%…+2.03% |
| PPT, 30 fixtures | two-point | `70,833 + 4.3409 × bytes` | 2 | 2.1704 | +0.60%, −1.94%…+2.17% |

The OLS row is the estimate; the two-point row interpolates through only the
smallest and largest fixture of the format and is reported beside it as the
check that the relation is linear rather than an artefact of the middle. The two
estimators differ by 0.005% on the DOC per-byte term and 0.18% on the PPT one.
Residuals are **(predicted − measured) / measured**, stated in that direction.

The two formats' OLS constants agree to **0.51%** across 38 fixtures and a 149×
size range. They reproduce 0609's fit (`217,274 + 12.90 × bytes`, i.e. 2.150 per
scan) to within 0.6% on the per-byte term at a different base commit, and 0589's
independently measured 2.048 (range 2.029–2.115) to within 5.6%. The predicted
saving of removing *k* scans is `k × 2.16331 × bytes` cycles — the DOC **OLS**
constant, unrounded, which is what `scripts/predict.py` and every table below
use — and it is **modelled** from a measured constant.

Per admitted DOC fixture, `SourceSnapshot::open`
([`perf/predicted-savings.txt`](results/change-0644/perf/predicted-savings.txt)):

| fixture | bytes | measured now | **B3+C** (6→4) | Δ | B2+C (6→3) | Δ | B1+C (6→2) | Δ |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `documentProperties.doc` | 9,728 | 332,564 | 290,475 | **−12.7%** | 269,430 | −19.0% | 248,385 | −25.3% |
| `endingnote.doc` | 9,728 | 332,802 | 290,713 | −12.6% | 269,668 | −19.0% | 248,623 | −25.3% |
| `footnote.doc` | 9,728 | 335,732 | 293,643 | −12.5% | 272,598 | −18.8% | 251,553 | −25.1% |
| `noheadfoot-litchi.doc` | 10,240 | 338,695 | 294,390 | −13.1% | 272,238 | −19.6% | 250,086 | −26.2% |
| `lists-margins.doc` | 10,752 | 346,910 | 300,390 | −13.4% | 277,130 | −20.1% | 253,870 | −26.8% |
| `table-merged-cells.doc` | 17,408 | 428,814 | 353,496 | −17.6% | 315,837 | −26.3% | 278,178 | −35.1% |
| `duplicate-style-names.doc` | 64,512 | 1,034,836 | 755,717 | **−27.0%** | 616,157 | −40.5% | 476,598 | −53.9% |
| `picture.doc` | 1,448,448 | 19,006,126 | 12,739,236 | **−33.0%** | 9,605,790 | −49.5% | 6,472,345 | −65.9% |

**B3+C, the recommendation: median −13.2%, range −33.0% to −12.5%** — Option B
does not touch the open under B3, so this column is also Option C's saving
alone. B2+C: median −19.9%, range −49.5% to −18.8%, 6.7 percentage points more
at the median for cost B-iii's second public item and error-branch re-hash.
B1+C: median −26.5%, range −65.9% to −25.1%, a further 6.6 points for cost
B-ii's wider trailing window. The scans removed are 2, 3 and 4 of six
respectively, priced at the unrounded OLS constant (12.979873 / 6 = 2.1633122);
recomputing from the four-figure value drifts by a few cycles in the last digit,
and from the two-point constant by about four. The readback gains the same shape again: `resolve` calls
`ensure_current` twice. Under B3 and B2 each `ensure_current` costs three
complete scans instead of four — its first call reduced, its last keeping its
pair — so `resolve` falls from eight to six and a paragraph read saves a further
`2 × 2.16331 × bytes`: **42,089** cycles on the smallest fixture and
**6,266,890** on `picture.doc`
([`perf/predicted-savings.txt`](results/change-0644/perf/predicted-savings.txt)).
Under B1 it falls to four, saving twice that.

**PPT: predicted change zero**, on all 30 fixtures. Option B does not apply
there and Option D is rejected, so the PPT open keeps both scans and both index
parses. This is stated as a prediction because the shared primitive is what
changes; the gate below exists to prove it.

**A/A floor, same window, same counters:** a second identical leg against the
first gives median **−0.15%** on cycles per operation over the 38 cells, range
−2.45% to +0.47%, 95th percentile of the absolute deltas **1.84%**
([`perf/perfstat-legB.tsv`](results/change-0644/perf/perfstat-legB.tsv)). That is
inside the host's standing p50 ≈ 4% figure, and the **smallest** predicted effect
(B3+C's −12.5%) is 83× the median floor and 6.8× the 95th-percentile floor.

**What the saving does not do.** 0609 concluded that the source-backed route
could not be made competitive with the eager `.doc` route by removing
fingerprint passes, and this design confirms it with the stronger reduction 0609
did not model. Route E was **not** remeasured here; its figures below are 0609's,
taken at its own base commit. On `noheadfoot-litchi.doc`, route E is 84,029
cycles and route S is 338,695 now (4.03×) and would be 294,390 under B3+C
(**3.50×**), 272,238 under B2+C (3.24×) or 250,086 under B1+C (**2.98×**); on
`documentProperties.doc`, 67,643 against 332,564 now (4.92×), 290,475 under B3+C
(**4.29×**) and 248,385 under B1+C (3.67×). 0609 modelled a 6→4 reduction at
3.6× on `noheadfoot-litchi.doc` and this record models the same variant at
3.50× from a measured before-value;
even the 6→2 variant only reaches 3.0× and stops there, because what remains is
the intercept — five, now four, index parses and the semantic parse — not the
per-byte term.
**This design improves the source-backed path; it does not change 0609's verdict
on the facade route, and this record does not reopen it.**

There is one saving the corpus makes visible that **the adopted design does not
take**, and it is the strongest argument an implementing change could make for
B2. 48 of the 57 `.doc` fixtures — **84% of the corpus** — are refused *after*
paying two complete scans, reading 2.0×–3.2× the artifact
([`counts/census-doc.tsv`](results/change-0644/counts/census-doc.tsv)). Those two
scans are identity call 1's pair: the refusal comes from
`reject_container_components` or `read_fib`, so call 2 and call 3 are never
reached and **Option C does nothing for them**. Under B2, which reduces call 1,
they would pay one scan and read about 1.0×–2.2× the artifact — a **31%–49%
reduction in bytes read** on 84% of the corpus, which the −19.9% open median does
not show because the eight admitted fixtures are the other 14%. Under B3 the
refusal path is unchanged. That is a reduction in scans and bytes, not a halving
of the refusal's cost, which also carries the index and range reads that do not
move.

## Admission gates for an implementing change

A change implementing B3+C is admitted only if all four pass. Where a gate's
number differs by variant it says so, and the implementing change must declare
which variant it took before the run rather than after it. The **before** half
of gates 3 and 4 is captured in this packet, so the implementing change measures
one leg.

**G1 — change-under-read sweep over every former read boundary, same typed
refusals.** Run the sweep of
`probe/main.rs` on `doc-open`, `doc-open-read` and `ppt-open`, over every read
ordinal of a clean run, on the 8 admitted `.doc` and all 30 `.ppt`, with the
witness byte placed (a) outside every parsed range and (b) inside the FAT sector
and (c) inside the directory sector. Required:

* **every ordinal refused today at or after the first complete scan is still
  refused**, from `open` or from the mandatory readback. This is the
  unconditional half of the gate;
* the **typed error** is unchanged at every ordinal, with the exceptions
  enumerated below **per witness placement** and no others, declared before the
  run rather than accepted after it. Under B3 the exceptions are Option C's and
  only Option C's, and they are the same six ordinals on every placement — 24–29
  on a 30-read open — because Option C deletes identity call 3, which is the
  reporting site for both bands. What those six report at the first access
  differs by placement, and the gate must say which it expects:
  * **payload witness** — `Overlay(SourceFingerprintChanged)`, from
    `ensure_current`'s comparison against the digest retained at open;
  * **FAT-sector witness** — `Overlay(Ole(CorruptedFile(..)))`, because the
    corrupted sector makes the first `ensure_current`'s composed reopen fail
    before any digest is compared. Rows 24–26 are I3c's `Overlay(Ole(..))` and
    27–29 are `R5 ≠ R6`'s `SourceFingerprintChanged` today; both become the
    former at the first access, and a gate that demanded
    `SourceFingerprintChanged` there would be wrong;
  * **directory-sector witness** — identical to the payload witness on
    `documentProperties.doc`, where the flip at offset 8,800 does not corrupt
    the chain; on a fixture where it does, the FAT-sector expectation applies;
* the remaining 25 ordinals of each sweep are **unchanged in variant and inner
  message**. In particular an implementation that reports
  `Ole(CorruptedFile(..))` where the packet's FAT sweep reports
  `Overlay(SourceFingerprintChanged)` — rows 7–9 — has taken B2 or B1 without
  cost B-iii's relocation and **fails this gate**; and one that reports
  `Overlay(SourceFingerprintChanged)` at rows 5–6, where the packet reports
  `Overlay(Ole(CorruptedFile))`, has relocated onto the composed reopen's own
  error branch and **also fails**;
* the trailing `OK` window is **one ordinal on `doc-open` and one on
  `doc-open-read`**, as it is today, because both B3 and B2 keep the last
  identity call of each operation and so end each in a compared scan. An
  implementation that took B1 reports **5 on both modes**, which is cost B-ii
  and must be argued in its own record rather than passed off as B3 or B2;
* **14** of W-D2's 17 `Overlay(Ole(..))` rows keep both wrapper and site —
  I3a's 2–6 and I3b's 12–20 — which is the test that Option D was not
  implemented by accident. The other three, I3c's 24–26, are covered by the
  Option C bullet above and must not simply disappear;
* the PPT sweep is **byte-identical** to
  [`sweeps/sweep-45543-ppt-open.tsv`](results/change-0644/sweeps/sweep-45543-ppt-open.tsv)
  and its index-region companion, which is the test that Option B did not leak
  into the PPT open;
* the trailing control — a mutation after the operation's last read — still
  returns `OK`, so the sweep cannot pass by refusing everything;
* the refusal **count** per fixture is asserted exactly, in the style 0589
  established, so a lost fence point fails rather than silently shrinking the
  count.

**G2 — byte-identical publication over the 57 DOC and 30 PPT fixtures.** Change
0589's 87-artifact differential, rerun unchanged: for every `.doc` and `.ppt`
under `test-data/`, both legs print the empty-splice identity plan's two
fingerprints and `is_noop`, the SHA-256 of the bytes it publishes, the publish
report's two fingerprints, an exact byte no-op splice and an effective one-byte
splice over the first non-empty stream with their span counts and published
digests, and the DOC and PPT snapshot fingerprints or the exact typed refusal.
The two reports must be **byte-identical**. A design that changes the number of
scans must change no digest and no published byte.

**G3 — deterministic read counts.** The recording trace of this packet, rerun.
Required on all 38 fixtures, per the declared variant: DOC `open` **4**
complete scans under B3 (3 under B2, 2 under B1, from 6 today); DOC
`open` + `paragraph(0)` **10** under B3 (9 under B2, 6 under B1, from 14
today — B3's open 4 plus `resolve`'s 6); PPT `open` **2** (unchanged in every
variant); index parses DOC 5 → 4 and PPT 2 → 2; and `read_bytes` matching the
arithmetic. On `documentProperties.doc` that is **24** read calls for the open
(from 30) and **51** for `open` plus `paragraph(0)` (from 59) under B3, with the
last read of each a complete scan.
`version()` and `len()` counts **fall** and must be reconciled rather than
asserted unchanged: on `documentProperties.doc` a DOC open takes 108 `version()`
and 32 `len()` calls at this base — change 0589's packet recorded 116 and 35 at
its own base, so an implementing change must recapture the before value rather
than reuse that one — and the only permitted reductions are those
attributable to the removed `fingerprints` calls and the removed identity call —
the five `ensure_source_identity` sites and the PPT counts must be untouched,
which change 0621's per-site attribution method can pin. This is the primary
gate: it is deterministic, and a count is what the design actually claims.

**G4 — native cycles on the eight admitted DOC fixtures and all 30 PPT
fixtures.** `perf stat -r 3` isolation pairs, same pin, same N, against
[`perf/perfstat-legA.tsv`](results/change-0644/perf/perfstat-legA.tsv), with an
A/A leg in the same window. Required: every DOC fixture within ±5% of the
predicted after-value for the **declared variant** in the table above, **no PPT
fixture adverse by more than 5%** against its before value (the floor here is
1.84% at p95), and every scenario reported including any that gets worse. A DOC
result that beats the declared variant's prediction by more than 5% is equally a
failure: at the measured per-scan constant it would mean a further scan was
removed, which is a different variant with different costs.

## Why it is sound

* **No fence is replaced by a weaker one.** Option A is rejected outright and
  every `ensure_source_identity`, `check_source_version` and `ensure_length`
  stays exactly where it is. The complete-artifact SHA-256 remains the only
  thing that can establish that two observations saw the same bytes.
* **Exactly two refusal boundaries move under B3+C, and both are stated.**
  Option C relocates six ordinals of the open from `open` to the first access;
  Option B, inside `ensure_current`, stops refusing the self-erasing transient
  witness W-T3 demonstrates there. Every *held* mutation from the first complete scan onward is still
  refused, and the trailing window is one ordinal on both measured modes before
  and after. Options D, E, F and G would each have moved a third boundary, and
  are rejected for that reason.
* **The error-precedence change B could have made is avoided, not merely
  repaired.** Removing an open identity call's confirming scan would hand
  W-D2's 7–9 window from `Overlay(SourceFingerprintChanged)` to
  `Ole(CorruptedFile(..))`, blaming the file for damage the caller's writer did,
  and repairing that by change 0621's relocation rule turns out to need a
  reopen-free re-hash on `source.rs:531`'s error branch — a second public item.
  B3 sidesteps the whole question by leaving the open's identity calls alone;
  inside `ensure_current` no retained index parse follows a reduced call, so
  there is no precedence to preserve. Gate G1 fails an implementation that takes
  B2 or B1 without the relocation, and equally one that relocates onto the
  composed reopen's own error branch.
* **ADR reading.** ADR 0003 makes the fingerprint diagnostic and exact byte
  equality the authorization; nothing here changes either. ADR 0005's *"mutation
  during a read returns `SourceChanged`"* holds at every ordinal the sweep
  covers except the two the bullets above name. ADR 0006's source identity is
  the same digest over the same complete artifact, compared at the first of the
  DOC open's own two comparison points (the `if` at `source.rs:593`); the second
  (`source.rs:601`) is what Option C removes. Proposed ADRs 0030 and 0031 are
  not accepted and are not cited.
* **Contracts untouched.** No limit relaxed, no `unsafe`, no dependency, no
  threading, no ambient I/O. **B3 needs exactly one public addition**: the
  empty-splice identity entry point in `litchi-cfb`, which takes one scan and
  keeps the composed reopen; taking no splices at all is what stops it being
  used to skip the reopen for a real plan. B2 and B1 need a **second** — a
  reopen-free digest for the error branch — and that is the single biggest
  reason they are priced rather than recommended.
* **The owned path is unaffected.** `open_owned_source` (`source.rs:637`)
  already takes one identity call on a source the module seals, and
  `source_is_owned_immutable` already elides its confirming scan
  (`overlay.rs:1057`). Option B gives the generic path the *caller-bracketed*
  version of the same reduction, on a different proof: ownership for the owned
  path, an outer comparison for the generic one.

## Correctness evidence

No production code changed, so there is no behaviour to test. What this record
asserts about behaviour was measured:

* **The read inventory** is traced, not inferred: 30 reads on
  `documentProperties.doc`, 46 on `picture.doc`, 8 on `45543.ppt` and 59 on
  `open` plus `paragraph(0)`, each read classified as a scan chunk or a parse
  read, with the complete-scan count derived as
  `scan_chunks / ceil(bytes / 1 MiB)` and equal to 6.00, 6.00, 2.00 and **14.00**
  exactly. The readback's 8 are `resolve`'s two `ensure_current` calls at two
  identity calls of two scans each, which is where the adopted design's Option B
  saving comes from in its entirety.
* **Every fence point** is located by a change-under-read sweep over every read
  ordinal, on three witness-byte placements, with the clean control and the
  trailing control both in the retained output. The sweeps were run on
  `documentProperties.doc` and `45543.ppt` — one fixture per format — not on the
  whole corpus; gate G1 requires the whole corpus of an implementing change.
* **Four transient witnesses** bound Option B's loss from both sides, twice:
  W-T1/W-T2 on the open's first identity call, which is where B2 and B1 would
  take it, and W-T3/W-T4 on `ensure_current`'s first call, which is where the
  adopted design does.
* **The `SourceVersion` witness** runs against an ordinary `FileSource` and a
  real filesystem, not against a synthetic adapter.
* **The witness for the adopted variant's own cost was run**, not argued by
  analogy: W-T3 and W-T4 exercise `ensure_current`'s first identity call, which
  is the only call B3 reduces.
* **The census** covers all 57 `.doc` fixtures with the bytes read before each
  refusal, which is what prices Option E.
* **The per-scan constant** is estimated independently on two formats, by OLS
  over every fixture of each, and the two agree to 0.51%; a two-point
  interpolation through each format's extremes agrees to 0.33% and is reported
  beside it as the linearity check.

### Gates

Run in `/home/zhuhe/code/litchi-worktrees/0644`; tails in
[`gates.txt`](results/change-0644/gates.txt).

| gate | result |
| --- | --- |
| `git diff --name-only c7326f680 -- crates/` | empty — no crate file changed |
| `diff -rq` this worktree's `crates/` against the shared before checkout's | empty — the trees are identical |
| `cargo fmt --all --check` | pass |
| `cargo doc -p litchi-cfb -p litchi-doc -p litchi-ppt --no-deps` | pass (rustdoc lints are `deny`, no warning emitted) |
| probe build against this worktree's `litchi-core`, `litchi-cfb`, `litchi-doc`, `litchi-ppt` | compiles clean, no warning |
| worktree-built probe against before-built probe, same sweep | byte-identical output |

Per-crate `clippy` and `test` gates are not run because no crate was touched;
`cargo doc` is run because this record's claims are about rustdoc-visible
contracts, and the probe's compilation is the evidence that the tree is intact.
No pre-existing failure was encountered. The harness suite in
`tools/perf-baseline` and the feature-bearing `cargo test -p litchi` are not run
because no crate and no harness file changed, and no registered selector reaches
either snapshot path (change 0587's measurement blocker, which this record does
not lift).

## Validation preserved

Everything, by construction: no reader, no fence, no limit and no refusal
changed, because nothing changed. The record's purpose is to state precisely
which of them an implementation may move and which it may not. Options A, D, E,
F and G are recorded as rejected so that a later reader does not rediscover
them; Options B3, B2, B1 and C are recorded with the witness for each detection they
give up.

## Limitations — what is not claimed

* **No speedup or regression is claimed and no claim-registry entry is made.**
  `performance_claim: none`. The predicted savings are **modelled** from a
  measured per-scan constant and a measured before-value; nothing was
  implemented, so no after leg exists.
* **Costs B-ii and B-iii are derived from the traced read order, not from a run
  of the changed code.** That removing an open identity call's confirming scan
  would hand the FAT-sector sweep's 7–9 window to `Ole(CorruptedFile(..))`; that
  relocating a scan onto the composed reopen's own error branch would invert
  rows 5–6; that B1 would grow the trailing window from 1 ordinal to 5 on both
  modes; and that a re-plan on `source.rs:531`'s error branch would fail on the
  same corrupted sector and yield a third variant — all four follow from the read
  order the trace measures and from where `finish_overlay_plan_with_owner` places
  the composed reopen (`overlay.rs:1035-1041`). None was produced by running a
  modified `litchi-cfb`. The reopen-free re-hash that would repair B-iii is a
  design, not a demonstration; gate G1 is what would prove it, and B3 is
  recommended partly because it does not depend on it. The argument that
  `ensure_current`'s own trailing reopen is harmless — that the paragraph bytes
  are read at r43–47, before either scan of the trailing call, and that
  `observed` is compared against the digest retained at open (`source.rs:768`) —
  is likewise read off the call order, and it is what grounds gate G3's figure
  of **10** scans for `open` plus `paragraph(0)` under B3.
* **The variant set is not exhaustive.** B applies per identity call, so other
  assignments exist; the three named here are the ones whose costs differ in
  kind — no reduction on the open at all, a reduction that needs the
  error-branch re-hash, and a reduction that also gives up the last-call
  exception. A fourth combination could apply B to the open but keep
  `ensure_current` whole; it was not priced because nothing recommends it.
* **Choosing B3 over B2 and B1 is a judgement, not a measurement.** No variant
  serves a wrong value; the reads after an operation's last scan belong to a
  composed reopen whose candidate is dropped. B3 is preferred because it is the
  only one that needs a single public addition and no error-branch machinery,
  and because `docs/GOAL.md` asks for the smallest coherent change; a different
  reader could reasonably take B2's further 6.7 percentage points and argue the
  second entry point in its own record. The record's job is to make that trade
  visible, not to foreclose it.
* **The per-scan constant is an estimate over 8 and 30 fixtures, not a claim
  about one.** The record reports OLS and a two-point interpolation side by side
  because they differ (0.005% on the DOC per-byte term, 0.18% on PPT) and the
  reader should see which one a number came from; the predicted-saving table
  uses the DOC two-point constant throughout.
* **Option C's relocation is modelled, not demonstrated end to end.** The
  receiving fence's behaviour is measured on the adjacent window (triggers
  30–58 of the readback sweep); that triggers 24–29 land in the same class under
  C follows from the read order, and only an implementation can prove it.
* **The sweeps cover one fixture per format.** The read *structure* is confirmed
  on a second DOC fixture (`picture.doc`, 46 reads, 6.00 scans, the same site
  order) by the trace, but no sweep was run on the other seven `.doc` or the
  other 29 `.ppt`. The trigger ranges quoted throughout are
  `documentProperties.doc`'s and `45543.ppt`'s; the *classes* they name are
  structural and should carry, which is what gate G1 is for.
* **The PPT conclusion rests on one fixture's sweep and on code reading.** The
  index-region witness was run on `45543.ppt` only; that PPT's `ensure_current`
  checks version and length alone is read from `text_edit.rs:515-530` and from
  0589, not measured by serving a mutated `read_text`.
* **No cold-cache, physical-device, range-source, remote, cross-platform,
  allocation or RSS result.** Every measurement is warm, local, single-threaded,
  on one host with SHA-NI, with eight agents building concurrently. A host
  without SHA-NI would make the per-scan constant larger and the predicted
  saving larger with it; none was measured. The `FileSource` witness is Unix
  only — the Windows policy tracks length, creation time and last-write time and
  would need its own witness.
* **The corpus ceiling is 1.45 MB** for DOC and 1.34 MB for PPT, with no
  DIFAT-scale artifact and no 4,096-byte-sector CFB. The per-byte term means the
  saving grows with size, so the corpus understates it; the fit is not
  extrapolated beyond the measured range in the table.
* **The census classification is by bytes read, not by call site.** A fixture is
  called "refused before the first complete scan" when it read less than one
  artifact; the attribution to I1 is read from the error variant, not from a
  backtrace.
* **The 2 GiB amplification figure for Option E is arithmetic**, not a measured
  denial-of-service demonstration. The measured half is the one fixture in the
  corpus and its 512 bytes.
* **The A/A floor stated here** (median −0.15%, p95 1.84%) is this window's
  floor on a pinned core; it is not a claim about the host in general.
* **Option B needs a public entry point in `litchi-cfb`** and this record does
  not design its name, its module or its error type — only its contract: empty
  splices only, one scan, the composed reopen retained, and a rustdoc statement
  that the caller owns the read-twice-compare.

## Retained evidence

[`results/change-0644/README.md`](results/change-0644/README.md) — the contents
table, provenance, the probe source and every raw output cited above.
