# Log sections for change 0620

Four paragraphs for the coordinator to merge, one per log, in the style of each
file's newest section. This change does not edit `HOTSPOTS.md`, `GOAL_AUDIT.md`,
`REPORT.md` or `ADR_COMPLIANCE.md` itself.

---

## For `HOTSPOTS.md`

## 0620 — the XLS edit-and-save path attributed; one duplicate source parse removed; the readback-owner swap frozen as inadmissible

Takes change 0587 item XLS-9 (rank 31), which the survey could only size as
**unknown** because no attribution of the XLS edit path existed, and supplies
it. On `54016.xls` one `Snapshot::from_bytes` is 167,311,992 Ir and the complete
eager `Workbook::new` inside it is **82.1%** of that; the offset inventory the
edit owner actually keeps is 12.5%. One `commit_source_backed` runs **three**
more complete parses of the same workbook — for the source policy checks, inside
the target's `Snapshot::from_bytes`, and inside
`verify_source_backed_numeric_target` — so **one open-edit-save parses the whole
workbook four times, 52.1% of the operation**. 0587's falsification test
("falsified if `Workbook::new` is under 15% of a plan-only commit") resolves at
**54.7%**; on the generic path it is 76-80% of the commit. One of the four
parses was already redundant: `commit_source_backed_numeric` re-parsed the
snapshot's own sealed bytes to ask three questions `SourcePolicyFacts` had
already answered at open, which commit `237309eea` had fixed for the plan-only
sibling and left here. Reusing the facts removes **137,061,668 Ir**, cuts the
operation by 14.0% in instructions and **25.1% in native cycles**, cuts
allocation calls by 19.9% and **peak retained bytes by 34.1%** (the removed
`Workbook` was alive across the whole target materialization), and moves the
commit p50 by **−32.1% / −33.7%** on `54016.xls` and −26.8% / −26.2% on
`WithCustomViews.xls` against an A/A floor under 2%. The remaining three parses
are named with their sizes. The survey's proposed cure — a source-backed
candidate readback — is **frozen as inadmissible as specified**:
`require_public_worksheet_coverage` asserts that the *eager parser* projected a
sheet, and `SourceBackedWorkbook` sets the same field unconditionally from the
BoundSheet8 `dt` byte without reading a worksheet byte, so it is always true;
five admission gates are recorded. Two new blockers for this area: **33 of 126 fixtures never reach
a snapshot at all**, 24 of them because a worksheet was not published by the
complete reader — including the flagship `ConditionalFormattingSamples.xls`,
which no XLS edit measurement in this program can ever use — and
the registered XLS numeric corpora have a Workbook stream that is 0.48% of the
archive against 78.9-94.2% on real files, so they cannot price anything in the
BIFF record path. OLE2/OOXML optimization remains active; ODF is deferred until
completion and iWork excluded. [Change and
limitations](0620-xls-edit-save-attribution.md); [retained
evidence](results/change-0620/README.md).

---

## For `GOAL_AUDIT.md`

## 0620 — the XLS edit-and-save path attributed; one duplicate source parse removed; the readback-owner swap frozen as inadmissible

GOAL step 1 (eliminate unnecessary work) in the narrowest admissible form: the
work removed is one complete parse of bytes the same call chain had already
parsed, under the same limits and the same compatibility profile, whose three
verdicts were already recorded. No step was skipped and no later step was
reached — no layout change, no algorithm substitution, no parallelism, no SIMD —
and no validation was moved, relaxed or deferred, because the candidate readback
this record is about is untouched: the target is still materialized, reopened
through a complete `Workbook::new` and validated again by a second one. The
decision rules are satisfied in order: before measurements from the shared
read-only checkout of the base, hypothesis and mechanism stated, smallest
coherent change, correctness and adversarial evidence, after measurements with
identical setup. The **measurement-blocker-is-work** rule was applied rather than
noted: 0587 recorded that no attribution of this path existed and that the
harness's XLS attribution binary supports only `open`, `list` and `one-cell`, so
this change wrote a probe that drives the real editor on real fixtures and
retains its source. The **instructions-rank-work, cycles-price-latency** rule
decided how the result is stated: callgrind reports −14.0% and native `perf stat`
−25.1% on the same operation, and the gap is exactly the software SHA-256
valgrind is forced into, which inflates the fingerprint terms the change does not
touch. Evidence tiers: **measured** for every per-call-site Ir figure (32
isolation pairs, 64 annotated profiles), the 32 native counter pairs, the 48
counter rows, the 5,760 timed operations across four rounds and the 3,546
corpus-differential rows; **modelled** for nothing — this record makes no
arithmetic prediction; **unknown** for the cost of the two remaining target
parses beyond their measured Ir, and for the shares on any workbook with a record
mix unlike the three measured. Host quiescence is established for the retained
timing run and is *not* assumed: the first timing run of this change was taken
while a corpus sweep occupied two other cores, its A/A floor reached 30% at p50,
and it was discarded and re-run rather than reported. The honest control band is
stated as ±5% at p50 — larger than the measured ±2% A/A floor — because the two
binaries differ in code layout and the unchanged `open` phase shows it. OLE2/OOXML
optimization remains active; ODF is deferred until completion and iWork excluded.
[Change and limitations](0620-xls-edit-save-attribution.md); [retained
evidence](results/change-0620/README.md).

---

## For `REPORT.md`

## 0620 — the XLS edit-and-save path attributed; one duplicate source parse removed; the readback-owner swap frozen as inadmissible

Two files changed in `litchi-xls` — one function body and three tests —
`performance_claim: none`. The changed scenario is `commit_source_backed` on a
fixed-width numeric edit: commit p50 **35,940,283 → 24,400,164 ns** on
`54016.xls` (−32.11% forward, −33.66% reverse, p95 −31.85%/−31.28%, p99
−31.82%/−31.49%) and **3,336,644 → 2,443,878 ns** on `WithCustomViews.xls`
(−26.76%/−26.19%), against a same-binary A/A floor of +1.70% and −0.20%.
Deterministic counters on `54016.xls`: allocation calls 201,886 → 161,751
(−19.88%), allocated bytes 93,192,484 → 76,204,234 (−18.23%), peak live bytes
39,572,027 → 26,082,823 (**−34.09%**); native cycles 228,627,848 → 171,248,047
(−25.10%). Controls: every counter of `open`, `number-plan`, `number-generic`,
`string-generic` and `noop-generic` is identical on both legs on all five
fixtures; instruction counts on those controls move −0.29% to +0.50%; their
commit p50s move −0.17% to +5.27%, inside the ±5% band the `open`-phase drift
establishes. **No comparison exceeds the +5% review trigger** except the
sub-microsecond `noop-generic` commit (210-615 ns), whose floors are as large as
its deltas and which is reported without evidential weight. The **registered
selectors cannot resolve the change and are reported saying so**:
`xls_numeric_source_backed_number_edit_save` gives −4.40% forward and −0.88%
reverse against a 6% eager-control drift, and the RK/MulRK case −2.66%/−2.74%,
because those corpora's Workbook stream is 0.48% and 0.82% of the archive while
the removed work is proportional to the Workbook stream. Correctness: 126
fixtures × five publication paths × three runs per leg — 591 rows per run and
3,546 rows compared, **0 refusal-text mismatches, 0 mismatches on every reported
field except the artifact digest, and 0 digest mismatches** on the 587 rows whose
digest is reproducible; all twelve registered selector cases keep their `output_sha256` across all
four rounds, including the Number `f8f37064…` and RK/MulRK `ddf5d5b8…` digests
change 0138 recorded. Gates: `cargo fmt --all --check`, `cargo clippy -p
litchi-xls --all-targets`, `cargo test -p litchi-xls` (72 binaries, 1,393 passed,
0 failed, 1 pre-existing ignored doctest) and `cargo doc -p litchi-xls --no-deps`,
all clean. **A pre-existing defect is reported, not fixed**: the generic
`Transaction::commit` publishes a CFB whose storage directory entry order comes
from a `HashSet` iteration (`litchi-cfb/src/writer/core.rs:977`), so two fixtures
publish different bytes on every process — reproduced three times on the
untouched before checkout, 4 of 591 rows on two fixtures, 250 differing bytes out
of 1,103,360, all inside UTF-16LE directory names. No cold-cache, physical-device,
range-source, RSS, concurrency-scaling, real-producer or cross-platform result is
claimed. OLE2/OOXML optimization remains active; ODF is deferred until completion
and iWork excluded. [Change and
limitations](0620-xls-edit-save-attribution.md); [retained
evidence](results/change-0620/README.md).

---

## For `ADR_COMPLIANCE.md`

## 0620 — the XLS edit-and-save path attributed; one duplicate source parse removed; the readback-owner swap frozen as inadmissible

ADR 0003's publication rule — "public format editors publish only after their
staged CRUD operation and typed readback succeed", with "the complete package
reopened under the retained limits" — is the governing text, and this change is
deliberately on the other side of it: what was removed is a re-validation of the
**source**, not of the candidate. The candidate readback is untouched, still runs
twice over the materialized target (`Snapshot::from_bytes` and
`verify_source_backed_numeric_target`), and still carries
`require_public_worksheet_coverage`, `require_unprotected_workbook`,
`require_macro_free_workbook`, `carry_fixed_numeric_inventory` and
`verify_public_numeric_readback` unchanged; change 0016's standing instruction
("target a different source of whole-workbook work, not remove either retained
validation layer") is the instruction followed. ADR 0006's validation contract is
satisfied by construction: the reused verdicts were produced by the same
`Workbook::new` over the same sealed `Arc<[u8]>` under the same default `Limits`
and `CompatibilityProfile::Strict`, `Inner` is constructed at exactly three sites
of which two compute the facts from the workbook they validate and the third is a
pure retag, and the two reachable refusals return character-identical strings
("protected or shared workbooks…", "protected worksheets…", "macro-bearing XLS
sources…"). The third verdict is unreachable: `SourcePolicyFacts::from_workbook`
propagates the coverage error with `?` before constructing the struct, so a
snapshot whose coverage failed does not exist, and the corresponding `require()`
branch is a dead defensive branch — exactly as it already is on the plan-only
path. Order and reachability are preserved; the only behavioural movement is that
a protected or macro-bearing source now refuses *sooner*, without first parsing
itself. No new `unsafe`, no weakened limit or malformed-input defence, no hidden
global Rayon pool, no ambient I/O, no public type or API change; source lineage,
overlay fingerprints, same-length splice validation, the artifact-length check,
the exact `Patch` and inverse, and the exact no-op fast path are all untouched.
**The frozen design is where the ADRs bite and is why nothing else was built**:
moving the candidate readback's owner from the eager `Workbook` to
`SourceBackedWorkbook` would need a proposed ADR, because ADR 0003 names the
complete reopen, and because `require_public_worksheet_coverage` cannot be
carried across — it asserts that the *eager parser's* per-sheet projection
succeeded, a fact set at one line (`workbook/package.rs:210`) and defeated by
roughly forty distinct per-sheet failures the loop swallows at `package.rs:217`,
whereas the lazy owner sets the same field unconditionally from the BoundSheet8
`dt` byte at `workbook/source.rs:2093` without reading a worksheet byte. The two
owners also disagree in the other direction (the eager parser slices each sheet
to the end of the whole stream with no upper bound and never validates the BOF
substream type; the lazy owner bounds each sheet, validates the BOF, and refuses
`FilePass` unconditionally), and their worksheet index numbering diverges the
moment any sheet fails eagerly. Five admission gates are recorded. Separately,
ADR 0006's "serialization is deterministic unless a `Clock`, actor identity, or
cryptographic RNG is explicitly supplied" is **reported as deviated** on the
generic XLS commit path, at `litchi-cfb/src/writer/core.rs:977`, pre-existing and
out of this change's scope. OLE2/OOXML optimization remains active; ODF is
deferred until completion and iWork excluded. [Change and
limitations](0620-xls-edit-save-attribution.md); [retained
evidence](results/change-0620/README.md).
