# Log sections for change 0624

Four paragraphs for the coordinator to merge, one per document, in the style of
each document's newest section. Each is written to stand alone. Their links are
relative to `docs/performance/`, where the four log documents live, not to this
packet directory.

## For `docs/performance/HOTSPOTS.md`

## 0624 — deflate is the largest term in a save, and SAVE-6's threshold is wrong in both directions

Frozen design plus the gate measurement item CORE-4 / SAVE-6 of change 0587
(rank 36) was waiting on; **no production change**. The item set a two-part
gate — two or more regenerated members of 256 KiB or more, and deflate above
20% of save wall time — and both parts are measured here for the first time.
**The timing part passes everywhere**: summing every `zlib_rs`, `flate2` and
`deflate` symbol under `perf record -e cycles:u`, deflate is **58.85%** of the
harness dense-wide one-percent save, **61.00%** of a one-cell edit on
`no_drawing_patriarch.xlsx`, **61.94%** of an authored 50-slide PPTX save,
**44.18%** of a 40-member `.rels` regeneration on
`ConditionalFormattingSamples.xlsx` and **24.19%** of a one-cell edit on the
same workbook; on the dense-wide save one symbol,
`zlib_rs::deflate::longest_match`, is 45.25% of the whole save against
`validate_authored_xml` at 23.98%. An independent wall-clock ratio of publish
p50 to deflate-only p50 agrees on all eight scenarios measured. **The size part
fails on every real file**: across 336 OOXML fixtures and four realistic edit
routes — one added relationship, a relationship on every related part, a
one-cell edit and a one-percent edit — **zero** saves regenerate two members of
256 KiB or more, the median edit regenerates two members totalling about 4 KB,
and the only scenario in the program that clears the threshold is
`xlsx_one_percent_commit_save` on the generated dense-wide corpus (2 ×
1,893,450 B). **And the threshold predicts the wrong sign**: the one real
fixture with a regenerated member over 256 KiB is `no_drawing_patriarch.xlsx`,
whose changed set is one 3.38 MB member and one 631-byte member, and it measures
**0.997× at widths 2, 4 and 8** — no width can help it — while the 40 × ~871 B
`.rels` set the rule excludes reaches **3.671× at width 4** and the authored
50-slide save's 161 members, averaging 1,550 bytes, reach **3.873×**. A controlled
crossover sweep at a fixed 871-byte member size finds profit from **four**
members upward (2.724× at width 4 on a 3.5 KB changed set), and a closed-form
bound `Σtᵢ / max(t₁, Σtᵢ/W)` predicts every measured cell within 8%: the
predictor is the *second*-largest member, not the largest. A/A floor in the same
window: serial p50 spread 0.21–1.04%, width-2 0.45–3.62%. Modelled end-to-end
(Amdahl over the measured section share and speedup, **not measured**): 1.85× on
the authored 50-slide save, 1.47× on the 40-member regeneration, 1.41× on the
dense-wide save, 1.00× on `no_drawing_patriarch.xlsx`. **Not implemented**, and
not for want of size: `docs/GOAL.md` rule 9 requires an explicit execution
context, the whole write path has zero `ExecutionContext`/`ExecutionLimits`
references, proposed ADR 0031 is not accepted, and change 0615 §7 orders gates
G1–G3 before any *use* of the parallelism — this record is that use. The
optimization order also puts a step-1 item ahead of it in the same profile
(`validate_authored_xml`, 23.98%). Next in this area: ADR 0031 accepted or
narrowed, then 0615's G1 and G3, then this record's gates A1–A6, and the
preservation writer's `PreservationIndex::prepare` before the streaming writer.
`performance_claim: none`; no claim is registered. OLE2 and OOXML remain active;
ODF is deferred until that goal completes and iWork is excluded.
[Change and limitations](0624-parallel-changed-member-deflate-design.md);
[retained evidence](results/change-0624/README.md).

## For `docs/performance/GOAL_AUDIT.md`

## 0624 — the audit's parallelism row now has a size, and its threshold has a correction

Design only; no production file was modified, so every limit, refusal, audit and
fence is exactly as change 0618 left it. Against `docs/GOAL.md` Workstream F's
"compression of changed output members" row, which change 0615 recorded as
having no parallel path at all, this record supplies the size the row lacked and
two corrections the audit should carry. First, the size: deflate of the changed
set is 24.19% to 61.94% of a save's native user cycles across five scenarios
spanning real fixtures, the harness corpora and the authored path, which makes
it the largest single term in every save measured. Second, the correction to the
gate: the "two or more members of 256 KiB or more" rule the survey inherited
from changes 0009, 0498 and 0499 does not transfer, because those records
measured *decompression batches* contending on one positional source with
per-wave thread creation and per-task memory reservations, whereas a deflate
task borrows an owned slice, shares nothing and returns a `Vec` — measured
profit begins at four 871-byte members, three orders of magnitude under the
proposed floor, and the real discriminator is whether one member dominates the
set. Third, the reason the row stays open: `docs/GOAL.md` rule 9 requires CPU
parallelism to be "opt-in and controlled by an explicit execution context with
thread, memory, I/O, cancellation, and task-granularity budgets", and a grep at
the base commit finds zero `ExecutionContext` or `ExecutionLimits` references
under `litchi-opc/src/pkgwriter.rs`, `atomic.rs`, the `soapberry-zip` writers or
the `litchi-cfb` writers; proposed ADR 0031, which would supply them, is not
accepted, and change 0615 §7 names its acceptance and gates G1–G3 as the
prerequisite for exactly this work. The audit rows this does **not** close:
nothing parallel was built and nothing end-to-end was timed, so every
end-to-end figure is Amdahl arithmetic and is labelled modelled; the DOCX
multi-part case the survey names was measured through the `OpcPackage`
publication route over all 62 DOCX fixtures rather than through `litchi-docx`'s
own editor, and the largest DOCX fixture's twelve ≥256 KiB members are embedded
fonts no paragraph edit regenerates; the authored path, which carries the
largest modelled win, sits behind the harder writer boundary and is out of scope
for a first implementation; and compression level remains fixed at 6 everywhere
and is still unmeasured, as SAVE-6 noted. Cold cache, peak RSS, allocation
profile, physical-device and cross-platform behaviour were not measured, and
allocation matters here because change 0618 rejected a pooled compressor on
glibc page-fault behaviour — this record makes re-measuring that an admission
gate (A5) rather than assuming per-worker state behaves differently.
`performance_claim: none`. OLE2 and OOXML remain active; ODF is deferred until
that goal completes and iWork is excluded.
[Change](0624-parallel-changed-member-deflate-design.md);
[evidence](results/change-0624/README.md).

## For `docs/performance/REPORT.md`

## 0624 — where a parallel deflate section would sit, and what it would have to prove

No production code changed. The design names one boundary and rejects the other
for a first implementation. The boundary it takes is
`PreservationIndex::prepare` (`crates/soapberry-zip/src/preserve.rs:840-914`),
which walks the plan's actions and calls `generated_entry(entry)` (`:897`, and
`:911` for appended members) for each `PreservationAction::Regenerate`.
`generated_entry` (`:2070`) is a pure function of one `&RegeneratedEntry`: it
compresses into a one-entry `ZipArchiveWriter`, re-parses that mini-archive for
its framing and returns a `PreparedEntry`, touching no shared state. Four
properties the code already has make the boundary low-risk: results are placed
at `prepared[index]` and emitted later by physical order (`:780`) and
central-directory order (`:808`), so the action list is explicitly not an
ordering control (`:236-243`) and reordering *when* a member is compressed
cannot move a byte; every compression already completes before any byte is
emitted, so the emission boundary does not move; `prepare` already retains every
regenerated member's buffer simultaneously — the fact change 0618's rejected
pooled variant turned on — so peak memory rises only by `W` compressor states;
and 0618's deliberate one-state-per-member choice becomes one-state-per-worker,
which has the opposite allocation shape and must be re-measured rather than
assumed. The boundary it **rejects for now** is `StreamingArchiveWriter`
(`office.rs:5744`, `write_deflated_with_accounting` `:6919`), which writes each
member into the sink as it compresses it: parallelising it requires a reorder
buffer, introduces in-flight bytes that do not exist today, and
`LimitedEntryWriter` charges the compressed-size budget per write call, so
buffering could move when a limit is refused — a contract change, and the place
the authored path's modelled 1.85× lives. Governance follows ADR 0031 §5-6 with
no new dependency edge: `soapberry-zip` defines its own `ScopedWorkers` exactly
as it already defines its own `CancellationProbe` (`office.rs:232-249`), a
`ParallelWriteSession` builds its pool lazily, `litchi_opc::OpenSession` bridges
`ExecutionLimits` to it as it already bridges `ParallelReadLimits`
(`litchi-opc/src/execution.rs:33-50`), cancellation is probed between waves and
never inside a member, and a new opt-in entry point carries the session so
`PackageWriter::to_bytes` and `save` keep their signatures and their serial
behaviour. The admission gates any implementation must clear are stated: byte
identity over 336 fixtures × five scenarios at four widths (A1); error identity
*and error order*, since the serial loop returns the first error in plan-action
order and a parallel version must return the same one rather than the first to
fail in time (A2); no regression at width 1 or below threshold, with no pool
constructed, counted from `/proc/self/task` (A3); scaling stated in full with
superlinear and `S < 1` cells labelled out-of-model (A4); minor faults and peak
RSS against change 0618's 271-fault finding (A5); and a fixture whose physical
local order differs from its central-directory order published at width 4 (A6).
Validation for this record is `cargo fmt --all --check` plus the repository's own
documentation and claim checkers; no crate was touched, so no crate suite
applies. See [Change 0624](0624-parallel-changed-member-deflate-design.md);
`performance_claim: none`.

## For `docs/performance/ADR_COMPLIANCE.md`

## 0624 — a design that adds no dependency edge, no visible executor and no ceiling

No production file was modified, so nothing in the accepted ADR set moves; what
follows is the reading the design commits to and the gates that would prove it.
ADR 0002 holds by construction: `soapberry-zip` gains no dependency, because it
defines its own `ScopedWorkers` trait rather than importing `litchi-core`'s,
exactly as it already defines its own `CancellationProbe` for the same reason,
and `litchi-opc` bridges the two as it already bridges `ExecutionLimits` to
`ParallelReadLimits`. ADR 0005's bounded resources hold: the parallel section
sits inside a preflight that already retains every regenerated member's buffer,
so peak memory rises only by `W` compressor states of about 300 KiB, `W` is
bounded by the granted worker permits, and the pool is built lazily on the first
qualifying batch — change 0615 gate G1's requirement, which `ParallelReadSession`
does not meet today because it builds its pool in its constructor. ADR 0006 is
the binding constraint and the design's entire claim to soundness is that it
cannot move a published byte: `generated_entry` is a pure function of one
regenerated entry, and output order is set by the index's physical local order
and central-directory order rather than by the loop, which the writer documents
as its contract. That is not argued but made an admission gate — 336 fixtures ×
five scenarios × four widths, byte-identical, with every typed refusal
reproduced — beside a second gate on error *order*, because the serial loop
returns the first error in plan-action order and a parallel version that
returned the first to fail in time would change error identity on a plan with
two failing members. ADR 0011 holds: `ParallelWriteSession` is a `soapberry-zip`
type reached only through an advanced-ingress session parameter, so no archive
type, raw lock or executor appears in any `Workbook`, `Document` or
`Presentation` signature, and `PackageWriter::to_bytes` and `save` are unchanged.
`docs/GOAL.md` rules 8 and 9 are the reason this is a design and not a change:
there is no ambient Rayon and no global pool today on the write path because
there is no parallelism at all, and adding any would require the explicit
execution context rule 9 mandates — which is proposed ADR 0031, **not accepted**,
and which change 0615 §7 places ahead of this work. Rule 10 is untouched: no
`unsafe` is proposed anywhere. This record also supplies a second, independent
argument for ADR 0031 §7's per-task floor, with the correction that a *deflate*
floor is on the order of a kilobyte rather than the 256 KiB the survey assumed,
and that the aggregate `min_parallel_bytes` cannot express the largest-member
rule the measurements actually support. See
[Change 0624](0624-parallel-changed-member-deflate-design.md);
`performance_claim: none`.
