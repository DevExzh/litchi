# Log paragraphs for change 0644

Four blocks, one per merged log. Each is written to be prepended as the newest
section of its file.

## For `HOTSPOTS.md`

## 0644 — the DOC snapshot's six complete reads are three pairs, and four of them turn out to be load-bearing

Change 0589 halved the *hashing* on the DOC and PPT source-backed opens and left
the *reads* where they were: a generic-`ReadAt` `SourceSnapshot::open` still
makes six complete artifact reads. A recording adapter now says exactly what
they are. On `documentProperties.doc` the open takes 30 reads, of which 6.00 are
complete sequential scans and the rest belong to **five** CFB index parses — I1
and I2 on the real source, plus one composed-CFB reopen inside each of the three
identity calls, which change 0589's inventory names but its read-count paragraph
does not total.
The six scans are three pairs, one per `identity_fingerprint` call, and the
comparison graph is `R1==R2`, `R3==R4`, `R5==R6` internally plus `R1==R3`,
`R3==R5` in `finish_open`. `picture.doc` has the identical structure at 46 reads
and `45543.ppt` at 8. The design that follows is **six scans to four** on the
DOC open and eight to six per `resolve`: the open loses only its third identity
call, which the trace shows brackets nothing — r24 and r25 are adjacent — while
the one-scan reduction is applied inside `ensure_current`, where the two calls
already bracket each other. It is kept *out* of the open because the first
call's confirming scan is the sole detector at one window of the FAT-sector
sweep, and repairing that precedence needs a reopen-free re-hash on
`source.rs:531`'s error branch — a second public entry point in `litchi-cfb`.
Predicted from a per-scan constant measured twice independently (OLS 2.1633
cycles/byte/scan over 8 DOC fixtures at 6 scans, 2.1743 over 30 PPT fixtures at
2, agreeing to 0.51% over a 149× size range), that is a median **−12.7%** of the
open's native cycles and **−33.0%** on `picture.doc`, against an in-window A/A
floor of −0.15% median and 1.84% at p95, plus a further `2 × 2.16331 × bytes` on
each paragraph read. Two more aggressive variants reach −19.9% and −26.5% at the
median and are recorded and priced, not recommended. The 84% of the `.doc` corpus that is *refused* is
untouched by the adopted design — those fixtures refuse inside identity call 1,
which B3 leaves alone — and that is the strongest argument for the B2 variant,
which would cut their bytes read by 31%–49%.
The queue item's own framing is adjudicated rather than reinterpreted: 0630
item 10 scopes this as "the third identity pass and the duplicate index parse,
2 of the 6 surviving complete reads", and the split is that the third identity
pass is admitted while the duplicate index parse is **rejected** — it is not a
complete read at all, and removing it would move the only pre-scan structural
refusal behind a complete hash. So item 10's two reads are exactly what the open gives up. Two further reads are
reachable from somewhere item 10 did not name — the confirming scan inside each
identity call — and both are recorded as priced variants rather than folded into
the headline; the place that scan comes out for free is the readback, which item
10 did not mention. Nothing is implemented; the record freezes the
design, the witnesses and the admission gates.
[Change and limitations](0644-ole2-snapshot-fence-design.md);
[evidence](results/change-0644/README.md).

## For `GOAL_AUDIT.md`

## 0644 — a source-identity fence priced per read, with the cheap-token substitution refused again

`docs/GOAL.md` puts correctness, lossless preservation and bounded resources
over speed, and ADR 0006's source identity is the rule this fence discharges.
This change moves the "finish source-backed CRUD adoption across formats" row
forward without touching it: the DOC snapshot's fence is now enumerated read by
read, each with the mutation it and no earlier read can detect, in the style
change 0621 established for XLS. The audit's most reusable result is a
*refusal*. Substituting a `SourceVersion` observation for any of the six
complete reads is rejected with a witness that runs on an ordinary `FileSource`
and a real filesystem: an in-place `pwrite` through a second descriptor with the
modification time restored by `File::set_times` leaves both `version()` and
`len()` unchanged while the bytes differ, which `FileVersionPolicy`'s own
rustdoc already warns of and which change 0587 classified as an ADR 0006 policy
amendment rather than an optimization. Three further options are refused for
reasons the audit should keep: dropping the composed-CFB reopen would rewrite
the `Overlay(Ole(..))` error wrapper for 17 of 31 sweep ordinals while buying an
estimated 0.55-12.89% of the open against the design's 12.5-33.0%, and unlike
B's confirming scan its precedence has no later observation to be relocated onto
at any price; hashing before
the first index parse would move the only pre-scan structural refusal in a
57-fixture corpus behind a complete hash, with a 2 GiB worst case on
attacker-chosen input; and deferring the identity in the style of change 0165's
lazy fingerprint would leave the semantic parse bracketed by nothing, which is a
wrong-value path rather than a relocated refusal. The one option admitted with a
cost — dropping the third identity call — moves six ordinals' refusal from
`open` to the first access with the same typed error, and the record names the
five accessors that consume no byte and would therefore return values for bytes
that may no longer exist. A second cost is **avoided rather than accepted**, and
that is the audit's other reusable result: removing an open identity call's
confirming scan hands one window of the FAT-sector sweep from
`Overlay(SourceFingerprintChanged)` to `Ole(CorruptedFile(..))`, blaming the file
for damage the caller's own writer did, and change 0621's relocation rule cannot
repair it with the identity entry point alone — the repair must sit on
`source.rs:531`'s error branch and needs a reopen-free re-hash, because a re-plan
there fails on the same corrupted sector and yields a third variant. That is a
second public item in `litchi-cfb` whose only job is an unbracketed digest, so
the design leaves the open's identity calls alone and applies the reduction only
where no index parse follows one. Change 0609's verdict on the facade `.doc`
route is **not** reopened: at the adopted design's minimum the source-backed
route is still 3.50×–4.29× the eager route on the two fixtures 0609 compared, and
2.98× even under the most aggressive variant. OLE2 and OOXML
remain the active priority; ODF stays deferred and iWork excluded.
[Change and limitations](0644-ole2-snapshot-fence-design.md);
[evidence](results/change-0644/README.md).

## For `REPORT.md`

## 0644 — one scan per identity call, where the caller owns the bracket

No crate file changed; `git diff --name-only c7326f680 -- crates/` is empty and
`crates/` is byte-identical to the shared before checkout. The design is: an
empty-splice-only identity entry point in `litchi-cfb` that performs the
planning scan and the composed reopen and returns the digest **without** the
confirming scan, applied inside DOC `ensure_current` — where the two calls
already bracket each other and no index parse follows a reduced one — but not to
the DOC open and not to PPT, plus dropping the DOC open's third identity call. Evidence
is deterministic first. The read trace gives 6.00 complete scans for DOC `open`
(30 reads on `documentProperties.doc`, 46 on `picture.doc`) and 2.00 for PPT
`open` (8 reads). Change-under-read sweeps place a held mutation after every
read ordinal on three witness placements: on DOC `open` triggers 0–4 return
`OK` (the mutation precedes the identity capture), 5–29 are refused with
`Overlay(SourceFingerprintChanged)` and trigger 30 returns `OK` as the trailing
control; over `open` plus `paragraph(0)` the run takes 59 reads and every
trigger from 30 to 58 is refused with the same variant by `ensure_current`,
which is the fence that receives the six relocated ordinals. Two flip-then-revert
witnesses, run on each of the two calls a variant might reduce, bound the one
detection the primitive change gives up from both sides: flipping after the
call's planning scan and reverting after its confirming scan is refused today
and leaves the artifact byte-identical, while reverting **one read earlier** is
already accepted today — on the open's first call (r5/r10 against r5/r9) and on
the first `ensure_current`'s, which is the one the adopted design reduces
(r31/r36 against r31/r35). Native `perf stat -r 3` isolation pairs over the 8 admitted
`.doc` and all 30 `.ppt` fixtures fit `206,301 + 12.979 × bytes` for DOC (6
scans; residual median +0.00%, −0.94%…+0.85%) and `68,254 + 4.3487 × bytes` for
PPT (2 scans), by OLS with residuals (predicted−measured)/measured of −0.26%
and −0.66% at the median, giving **2.1633** and **2.1743** cycles per byte per
scan — agreeing to 0.51%, and reproducing change 0609's fitted 12.90 per byte to
within 0.6% and change 0589's measured 2.048 to within 5.6%. A two-point
interpolation through each format's smallest and largest fixture gives 2.1632
and 2.1704, which is the linearity check. The predicted after-values are a
median **−12.7%** of the open (−12.5% to −33.0%, which is Option C's saving
alone because Option B does not touch the open), and a further
`2 × 2.16331 × bytes` per paragraph read, against an A/A floor of median −0.15%,
p95 1.84% over 38 cells. Two variants that do reduce the open are priced at
−19.9% and −26.5% and not recommended. **PPT is predicted to change by
zero**, because the primitive is not admissible there. No claim-registry entry;
`performance_claim: none`, and the savings are modelled from a measured constant
because no after leg exists.
[Change and limitations](0644-ole2-snapshot-fence-design.md);
[evidence](results/change-0644/README.md).

## For `ADR_COMPLIANCE.md`

## 0644 — the bracket is what makes the reduction sound, never the token

ADR 0005 requires that a mutation during a read return `SourceChanged`, ADR 0006
binds a splice to the bytes it was planned against, and ADR 0003 makes the
fingerprint diagnostic while exact byte equality authorizes application. The
design preserves all three, and it is precise about which one carries the
argument. The confirming scan inside each `identity_fingerprint` call may be
removed **only** because its comparison partner is an earlier scan the caller
already takes — inside `ensure_current`, `confirmed`'s partner is `observed` —
and not because the retained `ensure_source_identity` (a `version()`, a `len()`
and a `version()`) says anything stronger afterwards; that cheap check stays at
all five fence points of the DOC open and the three inside `ensure_current`, and
is proved insufficient by the `FileSource` witness. The same reasoning
**refuses** the change on the PPT open, whose single identity call has no outer
partner: the sweep shows four of its eight ordinals refused today by the
internal comparison and by nothing else, and PPT's `ensure_current` checks only
the version token and the length, so the loss there would be a `read_text` that
decodes bytes the retained identity does not cover, not a relocated refusal.
Error identity is treated as a contract rather than a detail: the FAT-sector
sweep resolves which parse reports each of the 30 ordinals and shows that the
composed reopens report `Error::Overlay(OverlayError::Ole(..))` while the two
real index parses report `Error::Ole(..)`, so the reopen is kept and the
admission gate requires that wrapper to be unchanged at every ordinal. I1 is
likewise kept in full rather than reduced to a header precheck, because it is
the malformed-input defence that bounds the work an attacker-chosen non-CFB
input can buy, and `docs/GOAL.md` forbids weakening one. The single contract
move the design accepts is stated rather than absorbed: dropping the third
identity call relocates six ordinals' refusal from `open` to the first call that
consumes a byte, and although every byte-yielding path goes through
`ensure_current`, the five retained-state accessors do not, so an implementing
change owes them a rustdoc sentence. No ADR is amended, no ADR clarification is
proposed, and proposed ADRs 0030 and 0031 are not cited.
[Change 0644](0644-ole2-snapshot-fence-design.md); `performance_claim: none`.
