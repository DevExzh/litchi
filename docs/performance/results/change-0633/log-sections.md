# Log sections for change 0633

Four paragraphs for the coordinator to merge, one per log, in the style of each
file's newest section. This change does not edit `HOTSPOTS.md`, `GOAL_AUDIT.md`,
`REPORT.md` or `ADR_COMPLIANCE.md` itself.

---

## For `HOTSPOTS.md`

## 0633 — the XLS commit's second complete target parse deleted; the open's two framing passes priced at 1.6% each and frozen

Takes the two items change [0620](0620-xls-edit-save-attribution.md) left on the
table for queue item **XLS-9** and resolves both, in opposite directions. **The
one 0620 could not remove is removed.** `commit_source_backed` materialized its
target, reopened it through `Snapshot::from_bytes` — which runs a complete
`Workbook::new` and then *drops it* — and then ran a **second** complete
`Workbook::new` over the identical bytes inside
`verify_source_backed_numeric_target`. 0620 recorded that removing it needed
"the snapshot retaining its reader", an ADR 0005 subsystem. It does not: the
verification moves *into* the target's own construction, where the first parse is
still a live local, and every check keeps its order and its position relative to
`parse_workbook_stream`, `Workbook::new`, `SourcePolicyFacts::from_workbook` and
`resolve_shared_strings` — so `require_public_worksheet_coverage` still runs
against the **source's** sheet list before `carry_fixed_numeric_inventory`
proves the two lists agree, which was 0620's exact blocker. The chain
`Workbook::new'verify_source_backed_numeric_target'commit_source_backed` is
**130,182,770 Ir before and 0 after** (12,109,922 → 0 on `WithCustomViews.xls`),
its drop glue 9,024,905 → 0, and the checks themselves cost the same
(4,419,288 → 4,420,759). One source-backed numeric commit falls **−15.71% in
instructions, −28.23% in native cycles, −24.80% in allocation calls and −22.46%
in allocated bytes**, and its `commit` p50 moves **−38.50% forward / −38.25%
reverse** on `54016.xls` and −34.88% / −36.75% on `WithCustomViews.xls` against
an A/A floor of −0.15% and +1.92%, on a host eight agents shared throughout —
the six p50 readings across three retained rounds span −38.1% to −44.1%. **The one the brief asked for is frozen,
because the mechanism is not where the cost is.** Reading the same annotations
down to their sub-1% rows prices the two framing passes of one
`Snapshot::from_bytes` at **2,579,172 and 2,579,104 Ir — 1.60% of the open each**
(0.92%/0.91% on `WithCustomViews.xls`), against `add_cell` at 33.82%,
`BTreeMap::insert` at 20.17% and the discarded `Workbook`'s drop glue at 4.68%:
both owners walk every record, only one pays for it. A new 29-case 0541-style
first-error matrix shows by row why the fusion is not value-identical — the
inventory frames first, so `cell XF index 4095 is outside 21 workbook resources`
shadows the eager parser's own XF validation, and the eager parser's `Err(_) =>
{}` per-worksheet arm turns two distinct defects into one generic
`worksheet at tab position 0 was not published by the complete XLS reader`. Two
smaller redundancies are named with their prices and frozen with the contract
question each raises: the Workbook stream is **extracted from the CFB twice**
(1.21% of an open) and the `LabelSst` text is **copied per cell by both owners**
(2.26%). One is implemented because it is value-identical as it stands: the
snapshot now retains the shared-string property table the reader already owns
behind an `Arc` instead of deep-copying it (2 to 58 allocation calls and 0.04% to
0.24% of allocated bytes per open, on four fixtures). Two findings for the area:
0620's CFB storage-order nondeterminism is **resolved at this base** by change
[0625](0625-cfb-writer-deterministic-storage-order.md) — all 591 corpus rows are
digest-reproducible where 0620 had to exclude four — and 0620's `analyze.py`
regex dropped every callgrind row under 10% because `callgrind_annotate`
right-aligns the percentage; the fixed copy is in this packet. OLE2/OOXML
optimization remains active; ODF is deferred until completion and iWork excluded.
[Change and limitations](0633-xls-commit-single-framing.md); [retained
evidence](results/change-0633/README.md).

---

## For `GOAL_AUDIT.md`

## 0633 — the XLS commit's second complete target parse deleted; the open's two framing passes priced at 1.6% each and frozen

GOAL step 1 (eliminate unnecessary work) in its narrowest form again: the work
removed is one complete parse of bytes the same call chain had just parsed, under
the same limits and the same options, whose result was still alive one stack
frame away. No later step was reached — no I/O restructuring, no layout change,
no algorithm substitution, no parallelism, no SIMD — and no validation was moved,
relaxed or deferred: the target is still materialized, reopened through
`PackageEditor::open` under the retained limits, inventoried, parsed in full,
checked by all five requirement checks, carried through
`carry_fixed_numeric_inventory` and read back by
`verify_public_numeric_readback`. ADR 0003's publication boundary and ADR 0006's
validation contract are untouched, and ADR 0005's cache contract is not engaged
because the snapshot retains no reader — the `Workbook` is a local binding inside
one constructor call, dropped before the `Snapshot` exists, and a plain open
still drops it at the statement it dropped at before. Change 0016's standing
instruction, "target a different source of whole-workbook work, not remove either
retained validation layer", is followed to the letter: both layers remain and the
checks' own cost is unchanged to within 1,471 Ir. The decision rules ran in
order: before measurements from the shared read-only checkout, hypothesis and
mechanism stated, smallest coherent change, correctness and adversarial evidence,
after measurements with identical setup. The **stop-at-a-design rule** did the
real work here: the brief asked for the two framing passes to be fused, the
before measurement showed the framing is 1.60% of an open rather than the 12.5%
and 82.1% the passes around it cost, and the first-error matrix showed neither
fusion direction preserves refusal identity — so part (a) is a frozen design with
its price and its admission conditions, and only its value-identical residue is
implemented. The **instructions-rank-work, cycles-price-latency** rule again
decided how the result is stated: callgrind reports −15.71% and native `perf stat`
−28.23% on the same operation, the gap being valgrind's software SHA-256
inflating the fingerprint terms the change does not touch. Evidence tiers:
**measured** for every per-call-site Ir figure (32 isolation pairs, 64 annotated
profiles), the 32 native counter pairs, the 20 (fixture, operation) counter rows
with four metrics each, the 11,520 timed operations across two complete
A1 B1 B2 A2 rounds, the 3,546 corpus-differential rows and the 58 matrix rows; **modelled** for nothing — this record makes no
arithmetic prediction; **unknown** for the cost of the two frozen redundancies
beyond their measured Ir, and for the shares on any workbook with a record mix
unlike the two measured. Host quiescence is **not** claimed: seven other agents
shared the machine, the first timing round's A/A floor on `54016.xls` reached
+12.3% at p50, and rather than hide it the capture was repeated twice, all three
rounds are retained, and only the round whose floors on the changed scenarios are
−0.15% and +1.92% is quoted at p50. The record states plainly that the unchanged
`54016.xls` controls cover ±15% across those rounds with flat instruction counts,
so no control timing on that fixture is a result in either direction.
OLE2/OOXML optimization remains active; ODF is deferred until completion and
iWork excluded. [Change and
limitations](0633-xls-commit-single-framing.md); [retained
evidence](results/change-0633/README.md).

---

## For `REPORT.md`

## 0633 — the XLS commit's second complete target parse deleted; the open's two framing passes priced at 1.6% each and frozen

Three files changed in `litchi-xls` — one constructor split in two, one private
struct, one helper, one deleted function, one `pub(crate)` accessor and three
tests — `performance_claim: none`. The changed scenario is `commit_source_backed`
on a fixed-width numeric edit: commit p50 **25,781,630 → 15,856,600 ns** on
`54016.xls` (−38.50% forward, −38.25% reverse, p95 −39.48%/−41.33%, p99
−38.49%/−40.70%) and **2,432,018 → 1,583,743 ns** on `WithCustomViews.xls`
(−34.88%/−36.75%), against same-binary floors of −0.15%/+0.25% and
+1.92%/−1.01%. Deterministic counters on `54016.xls`: allocation calls 161,751 →
121,641 (−24.80%), allocated bytes 76,204,234 → 59,090,612 (−22.46%), peak live
bytes 26,082,823 → 25,628,658 (−1.74%); callgrind 887,985,767 → 748,473,025 Ir
(−15.71%) and native cycles 179,110,791 → 128,547,808 (**−28.23%**). Controls:
`peak_live_bytes`, `published_bytes` and the complete source-backed diagnostics
object are identical on all 20 (fixture, operation) rows; the `open`,
`number-plan`, `number-generic`, `string-generic` and `noop-generic` controls
move −0.41% to +0.98% in instructions and −4.55% to +3.35% in native cycles.
**One comparison exceeded the +5% review trigger and was chased rather than
explained away**: the three `xls_owned_source_open*` registered selectors moved
+7.06% to +9.07% forward and +4.35% to +8.30% reverse at p50 against floors under
2.5%, on a path this change does not touch, and a callgrind isolation pair prices
one of their operations at **3,978,024 Ir before and 3,978,917 Ir after, +0.02%**
— identical work, host noise. The unchanged `54016.xls` commit controls likewise
cover −4.29% to +24.89% at p50 across the three retained timing rounds with
same-binary floors from −10.4% to +11.9% and instruction counts flat to within
0.5%, and the record says plainly that none of them is a result in either
direction. The single instruction-count regression is
`FormulaEvalTestData` at +0.38% to +0.98%, entirely inside the renamed
constructor's own inlined body with no named callee moving by more than 20,000
Ir, and it does not reproduce natively (−0.28% instructions). The **registered
selectors again cannot resolve the change and are reported saying so**: the two
source-backed cases move +0.32%/−1.86% and −1.94%/−1.84% on a function that moves
35-44% on a real workbook, because their corpora's Workbook stream is 0.48% and
0.82% of the archive — the second independent confirmation of change 0601's
open brief. Correctness: 126 fixtures × five publication paths × three runs per
leg — 591 rows per run and **3,546 rows compared, 0 refusal-text mismatches, 0
mismatches on every reported field except the digest, and 0 digest mismatches
over all 591 rows with none excluded** (0620 had to exclude four; change 0625
fixed the CFB storage order and sixteen runs per leg now yield one digest each).
A new 29-case 0541-style first-error matrix over synthetic malformed Workbook
streams is **identical on both legs**, and all twelve registered selector cases
keep their `output_sha256` across all four rounds, including the Number
`f8f37064…` and RK/MulRK `ddf5d5b8…` digests change 0138 recorded. Gates: `cargo
fmt --all --check`, `cargo clippy -p litchi-xls --all-targets`, `cargo test -p
litchi-xls` (72 suites, 1,402 tests), `cargo doc -p litchi-xls --no-deps`, `cargo
test -p litchi --features docx,xlsx,pptx,xls` (26 suites, 266 tests) and the
harness's own `cargo test` in `tools/perf-baseline` (19 suites, 531 tests), all
clean. **A defect in this program's own tooling is reported and fixed in the
packet copy**: change 0620's `analyze.py` matched `\(([\d.]+)%\)` while
`callgrind_annotate` right-aligns the percentage, so every retained row under 10%
was silently dropped; nothing 0620 reported depended on those rows, and
everything this record's framing attribution says does. No cold-cache,
physical-device, range-source, RSS, concurrency-scaling, real-producer or
cross-platform result is claimed. [Change and
limitations](0633-xls-commit-single-framing.md); [retained
evidence](results/change-0633/README.md).

---

## For `ADR_COMPLIANCE.md`

## 0633 — the XLS commit's second complete target parse deleted; the open's two framing passes priced at 1.6% each and frozen

**ADR 0003 (publication boundary).** "Publish only after their staged CRUD
operation and typed readback succeed" and "the complete package is reopened under
the retained limits" both continue to hold exactly. The materialized target is
still opened through `PackageEditor::open` with `Targets::default()` and
`Limits::default()`, still inventoried by `parse_workbook_stream`, still parsed
in full by an independent `Workbook::new`, still proved byte-identical to the
plan, still checked by `require_public_worksheet_coverage` against the source's
sheet list, `require_unprotected_workbook` and `require_macro_free_workbook`,
still carried through `carry_fixed_numeric_inventory` and still read back by
`verify_public_numeric_readback`. What is deleted is a *duplicate* of the reopen,
not the reopen: the second `Workbook::new` read the same `Arc<[u8]>` under the
same default limits and the same `OpenOptions`, and the first check of the moved
verification proves those bytes are the materialized target before any check
consumes the parse. The reading that "reopened" means "parsed twice" is the one
this change declines; the reading that it means "parsed independently of the
editor's own inventory, in full, under the retained limits" is preserved. The
neighbouring temptation — handing the eager reader the Workbook stream bytes
`PackageEditor` already captured, worth 1.21% of an open — is **frozen precisely
because it would make that second reading false**, and is recorded with the ADR
amendment it would need.

**ADR 0005 (cache contract).** Not engaged. Change 0620 judged that removing this
parse required "the snapshot retaining its reader", which would have made it an
ADR 0005 subsystem; it does not. The `Workbook` is a local binding inside one
constructor call, `drop`ped before the `Snapshot` is constructed, never stored in
`Inner`, never reachable from a public type, and a plain open still drops it at
the exact statement it dropped at before this change. No cache, no
lifetime-extended handle, no invalidation question.

**ADR 0006 (validation and determinism).** No validation was moved, relaxed,
reordered or deferred; the five checks keep their order and their position in the
construction, so no refusal moves. The 29-case first-error matrix is the
adversarial evidence for that and is identical on both legs; the 126-fixture
differential adds 591 typed-refusal rows and 591 artifact digests per run, all
identical. The determinism clause is better served than before: 0620 reported a
deviation — the generic commit's CFB storage directory ordered by `HashSet`
iteration — and this record confirms change 0625 closed it, with sixteen runs per
leg over the two affected fixtures now yielding one digest each and zero corpus
rows excluded as nonreproducible.

**No ADR was cited as authority for a change it does not cover, and proposed ADRs
0030 and 0031 are not cited at all.** No new `unsafe`, no weakened limit or
malformed-input defence, no hidden global Rayon pool, no ambient I/O, no public
leakage of archive types, raw locks or executors, and no public API, error type
or output byte changed. [Change and
limitations](0633-xls-commit-single-framing.md); [retained
evidence](results/change-0633/README.md).
