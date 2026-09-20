# 0704 — bounded PPTX slide MCE retention

Disposition: retain with scoped performance and memory costs. `performance_claim: none`.
Baseline revision: `d48523eec2`.

The [0703 diagnostic](0703-pptx-capture-projection-reuse-diagnostic.md)
identified repeated default MCE transformations across capture and changed
commit. This candidate retains only successful owned slide projections behind
an explicit opened-presentation policy. The default ceiling is 1 MiB; zero
disables retention. Marker-free borrowed projections create no memo entry.

## Candidate and constraints

An immutable snapshot owns the optional table. Raw allocation identity and a
strong raw owner establish reuse, while the existing package digest pass
supplies verified owners without new `Part::blob()` or `blob_arc()` observations.
Each hit still validates source/output limits, root, name, notes proof and the
surrounding relationships and identity constraints. No errors are memoized.
Snapshot, transaction and commit owners expose retained charges and release.

ADR 0005's explicit retention policy applies: output vector capacity and memo
metadata are charged, optional admission failure falls back to recomputation,
and release preserves semantic values and bytes. Per-value charges are not
additive physical memory when clones share allocations. The implementation
uses existing infallible `Arc::new` allocation behavior; logical accounting is
neither exact allocator RSS nor a guarantee of global OOM recovery. Shared
OOXML and OPC sources remain outside this candidate's edit boundary.

## Measurement scope

The [packet](results/change-0704/README.md) preserves a fresh 602-file baseline,
33 goal/accepted-ADR hashes and five build inputs. Thirteen native workflows
cover capture, clone, edit, commit, apply and total on prepared packages. The
initial two A/A legs have 100 samples and five warmups each. Whole-workflow
median drift is −1.66% to +1.74%. Allocation measurements use a separate
instrumented executable; ten refusal cases use their own bounded matrix.

The native profile covers the prefix through changed commit, including open,
capture, working clone, edit, commit and destruction. It excludes publication.
Counters at 10 and 210 iterations support a 200-iteration difference; this is
a separate denominator from the prepared-package workflow timings.

A candidate-only observer compares serialized patches, semantic digests,
revisions and published target text across default, disabled and tiny budgets.
It records charges through clone, edit, commit, publication and release, with
two fresh repeats on the real and generated sources. Both sources pass all parity and fresh-repeat checks. The real capture
charge is 539,340 bytes (538,356 output capacity plus 984 metadata); the
two-edit committed and published charge is 539,056 bytes. A released clone
reports zero while its source retains 539,340. Generated marker-free stages
report zero under every budget.

## Work-removal trace

Twelve fresh instrumented processes confirm the changed-commit mechanism.
Capture remains 18 default-MCE calls for the real deck and 19 for generated
slides. Real changed commit calls fall from the prior diagnostic's 18 to 6
for one edit and 7 for two edits: twelve and eleven unchanged slide
transformations are skipped. Generated changed commits remain at 19; both
no-op commits remain at zero. Setup and semantic verification have separate
phase labels and are excluded. Every call succeeds and both repeats agree.
The [trace audit](results/change-0704/mechanism/audit.py) checks exact source
restoration, raw/output digests and call counts. The temporary instrumented
executable supplies no timing evidence.

## Initial measured results

Initial ABBA results use 100 samples and five warmups per leg. The real
one-edit workflow improves by 18.88% / 18.69%, with changed commit improving
41.30% / 41.01%. Real two-edit workflow medians improve 16.18% / 18.50%.
Real capture medians increase from 4.5811 / 4.5607 ms to 4.6278 / 4.6112 ms.
The no-op real case costs 6.43% in the first candidate leg and 0.21% in the
second. This process-leg difference is preserved; its cause is not established.

All 13 workflow results and per-phase values are in
[tables.md](results/change-0704/tables.md). There are 73 native metric/phase
review triggers above 5%, chiefly short phases and tails, and two refusal p99
triggers: early-name +1,291 ns and late-root +5,220 ns in the second pair.
Longer follow-up measurements remain pending. No raw sample is discarded.

The real one-edit allocation companion reports 143,848 → 124,776 allocation
calls and 12,430,874 → 10,123,638 requested bytes. Peak live bytes above the
workflow start rise 459,170 → 954,239, and net live change rises
195,730 → 769,444. Capture alone retains 91,008 → 630,348 bytes, exactly the
539,340-byte observed logical memo charge on this build. The generated
one-edit control has identical allocation observables. All three allocation
repeats agree exactly. These are instrumented resource measurements, not
native latency samples, and the finite retention ceiling does not bound total
transient workspace or process RSS.

For the separately measured open-through-commit prefix, counter differences
per iteration report 224.42 → 179.54 million instructions and 52.61 → 43.26
million cycles. Whole-child peak RSS is 5,832 → 6,168 KiB. The candidate
page-fault difference is slightly negative (−0.015/iteration), illustrating
why two-process counter subtraction is diagnostic rather than an exact
per-operation resource guarantee.

## Longer follow-up and disposition

The full 13-workflow and ten-refusal matrices repeat with four ABBA legs,
300 samples and ten warmups: 15,600 native workflows and 12,000 refusal
captures. The [independent follow-up audit](results/change-0704/audit_followup.py)
recomputes all 264 phase/case comparisons from retained raw output.
[Follow-up tables](results/change-0704/followup-tables.md) show every workflow
and refusal case.

Real one-edit p50 improves 18.95% / 19.19%; two-edit improves 18.98% / 17.96%.
Real no-op p50 changes −4.02% / +0.10%, so the initial +6.43% leg does not
repeat. This does not erase the original result or establish its cause.
Follow-up early-name refusal p50 costs 5.57% / 2.56% (about 1.03 / 0.48 µs).
Marked late-missing-relationship p99 costs 46.98% (+67.18 µs) in the first
pair and improves 1.95% in the second. Notes-invalid-tail p99 costs 11.48%
(+23.63 µs) in the first pair and improves 3.08% in the second. Generated
no-op total p99 costs 5.48% (+37.54 µs) in the first pair. All 62 follow-up
triggers, including maxima and short-phase changes, remain in the packet;
there is no blanket noise attribution or tail-latency improvement claim.

Retain the candidate for the measured opened-PPTX edit path. Avoiding eleven
or twelve unchanged slide transformations yields a repeated roughly 18–19%
whole-workflow gain and fewer allocations/instructions. This is purchased
with about 539 KB of logical retained state on the real deck, higher live
memory and small capture/refusal costs. The default is finite and observable;
callers with tighter memory needs can set zero or explicitly release the
memo. Marker-free controls retain no projections and have unchanged measured
allocation counts. This is not a general document-open, save, streaming,
parallel, remote-I/O or cross-format speedup claim.

## Provenance and verification

The first baseline native compile failed after observing transient literal
patch text in a production file. The exact transient file was not captured.
The failed compiler log and patch-preparation artifact are preserved under
`results/change-0704/preflight/`. All 602 baseline hashes were restored and
all three executables rebuilt with before/after source checks before any
baseline timing used here. No failed-attempt executable or timing contributes
to this comparison.

All 26 focused tests and seven integration gates pass: formatting, all-feature
checks, warning-denied Clippy, default tests (1,248 passing results), full-feature
tests (4,172), facade tests (45), and warning-denied rustdoc. These totals include
repeated tests across configurations; they are not unique-test counts. The
recorded ignored tests remain ignored. Two independent source reviews accept
the final ownership/accounting and append/sort implementation. The final-table
reservation test forces capacity overflow and proves optional fallback releases
both candidate owners; pending-vector failure remains covered by source review
and ordinary admission tests. No fuzz campaign or cross-platform run is claimed.

All six evidence gates pass: crate boundaries, strict performance claims,
structural claims, report classification, CRUD coverage and the non-iWork
boundary. Cleanup and terminal sealing receipts record their final status.
