# 0525 XLSX unchanged-cell readback reconstruction (accepted)

This bundle documents the accepted source-backed OLE2/OOXML performance
experiment. It tests whether source-bound omission of unchanged XLSX cell owners can reduce
semantic reconstruction after a value-only edit while preserving independent
readback of the actual emitted XML. The candidate has been measured and passes
the frozen native and profile admission gates. The accepted
[`decision.json`](decision.json) retains the production change, and the
supplemental four-child eager confirmation also passes. Quality checks have 12
passing gates and 1,293 successful executions. Cleanup and strict post-cleanup
verification are complete, including exact report replay and absent owned
paths. Recursive seal verification also passes. Baseline captures include
the retained native A2
control. OLE2 and OOXML remain the active priority; ODF is deferred until that
optimization goal completes, and iWork is excluded.

The candidate is deliberately scoped to the existing source-backed XLSX
value-only path. The normal writer's complete worksheet bytes remain the
published output. A private proof records copied source cell spans and exact
owners. Complete output validation runs first. An eligible reduced document
retains worksheet context, row shells, changed cells and complete
membership-changing rows; the existing raw worksheet parser independently
parses that actual output. A fallible Store merge imports only proven omitted
source entries and rebuilds indexes and extents. New rows, membership changes
that cannot be isolated, shared-formula worksheets, unsupported forms, proof
mismatches, reduced-parse/merge failures and reservation failures fall back to
the complete parser. The reduced scratch is dropped before fallback.

## Why this candidate exists

The retained 0522 source-bound attribution puts the reconstruction branch at
about 61% of selected `MultiSourceEdit::commit` instructions and the rewrite
branch at about 39%, with disjoint attribution rows. Medium and dense-sparse
primary shapes contain 9,216 and 17,792 cells. A row-only design would omit
51.56% and 5.75% of cells respectively, while source-proven cell omission
could omit 98.99% and 99.00%. The exact values and the four baseline
attribution rows are in
[`mechanism-plan.md`](mechanism-plan.md) and
[`closure-coverage.json`](closure-coverage.json). They are design evidence,
not candidate timing results.

## Measured result

The native ABBA matrix retains 1,640 matched durations, with 820 per stage.
The original eager guard adds 240 samples across its four rows. The four
primary rows pass the frozen 5% total and commit p50 gates:

| shape / repeat | total p50 baseline → candidate (ns) | reduction | commit p50 baseline → candidate (ns) | reduction |
| --- | ---: | ---: | ---: | ---: |
| medium / 1 | 25,504,253 → 22,637,341 | 11.2409% | 11,574,154 → 8,020,944 | 30.6995% |
| medium / 2 | 25,534,145 → 22,298,803 | 12.6706% | 11,590,438 → 8,071,393 | 30.3616% |
| dense-sparse / 1 | 49,082,428 → 42,600,330 | 13.2066% | 21,715,666 → 15,180,594 | 30.0938% |
| dense-sparse / 2 | 50,541,246 → 43,010,549 | 14.9001% | 22,583,895 → 15,634,883 | 30.7698% |

The separate allocator lane retains 40 matched samples. Medium allocation calls,
allocated bytes and incremental region peak move from 118,744, 20,344,427 and
2,984,983 to 91,391, 12,436,307 and 2,344,427, reductions of 23.0353%,
38.8712% and 21.4593%. Dense-sparse moves from 225,771, 27,331,029 and
7,335,225 to 172,946, 18,258,492 and 4,467,508, reductions of 23.3976%,
33.1950% and 39.0951%. The vectors are identical across the two repeats and
their deallocation/reallocation balances reconcile.

The four shape/repeat profile pairs (eight isolated profile children) pass the
independent 15% instruction gate:

| shape / repeat | commit Ir baseline → candidate | reduction |
| --- | ---: | ---: |
| medium / 1 | 197,188,676 → 133,217,968 | 32.4414% |
| medium / 2 | 197,130,975 → 133,250,185 | 32.4053% |
| dense-sparse / 1 | 377,140,022 → 258,175,362 | 31.5439% |
| dense-sparse / 2 | 377,135,732 → 258,155,623 | 31.5484% |

The profile's raw worksheet parser inclusive Ir falls 96.0266–96.5004%, while
the complete validator remains within 0.019% of baseline. These are scoped
Callgrind diagnostics. The main comparison retains 31 over-5% adverse rows: 7
open-phase, 19 publication-phase, 4 reopen-phase and 1 whole-child-RSS row.
There are no over-5% main commit or total-elapsed rows. The RSS row is the
managed dense-sparse repeat-2 guard, where whole-child max RSS rises from
85,648 to 90,244 KiB (+5.3661%). Reopen diagnostics are outside the measured
edit/save elapsed time by the harness. The accepted decision retains these
observed phase/RSS tradeoffs only alongside total p50 gains in all four
primary rows; the evidence makes no blanket no-regression or RSS-improvement
claim. The comparison also retains 76 same-build drift rows. The original
eager guard retains one adverse metric: dense-sparse repeat 2 rises 4.5442% at
p50, 4.4781% at the mean, 4.7339% at p95 and 5.8175% at p99; it has no
primary-gain claim. The supplemental eager confirmation contains four children
(A1/B1/B2/A2), 20 warmups and 100 samples per child, or 400 samples total. Its
confirmation gate passes: A1/B1 candidate change is −2.543162% at p50 and
−2.622600% at the mean; A2/B2 is −3.799717% at p50 and −3.695763% at the
mean. All p95,
p99 and whole-child RSS comparisons remain below the adverse threshold, with
zero >5% matched or same-build drift flags. This bounded diagnostic does not
erase the original eager row shift or establish general eager acceleration or
host stability. The original eager rows remain retained.

The follow-up [`next-priority-review.md`](next-priority-review.md) selects the
`scan_with_limit` layout/provenance pass as the primary next source audit at
53.26% of these XLSX commit profiles' candidate commit Ir. The
`Store::merge_omitted_cells` path is the secondary audit at 5.97%. These are
profile shares for source review, not a measured opportunity or permission to
edit source.

The full contract and boundaries are described in the
[mechanism plan](mechanism-plan.md), the [reconstruction review](reconstruction-review.md),
the [adversarial review](adversarial-review.md), and the
[preflight corrections](preflight-corrections.md). The corrections include
the rectangular-bound merge fix, linear omission lookup, reduced-state
lifetime, existing-layout rendering, recoverable optional allocation and
validator marker-scan cleanup.

## Frozen campaign

The authoritative [plan](plan.json) and [driver](run.py) were frozen before
candidate build/capture. The primary matrix is:

| lane | frozen scope |
| --- | --- |
| native | medium and dense-sparse source-backed one-percent edit/save; 2 repeats, 20 warmups, 100 samples; primary total and commit p50 |
| guards | one-edit and managed one-percent medium/dense-sparse, vendor-extension, and noncompact/managed-noncompact; 2 repeats, 10 warmups, 30 samples |
| allocator | separate allocator build; medium and dense-sparse, 2 repeats, 5 samples; staged sets and commit only |
| Callgrind | isolated `MultiSourceEdit::commit`, 2 repeats, measured commit dump `.4` after three preceding lifecycle dumps `.1`–`.3` |
| hardware | whole fresh child, medium and dense-sparse, 2 repeats, 100 samples; diagnostic counters only |
| eager guard | ordinary rewrite consumer, medium and dense-sparse, 2 repeats, 10 warmups, 30 samples; no primary gain claim |
| eager confirmation | dense-sparse ordinary rewrite consumer; A1/B1/B2/A2, 20 warmups, 100 samples per child; 4 children / 400 samples; guard only |

Native execution is serial ABBA: baseline r1, candidate r1, candidate r2,
retained baseline r2 under the candidate checkout. The allocator lane reports
allocation calls/bytes and incremental region peak separately from native
latency. Callgrind's 15% commit-Ir requirement is independent of native
admission. Hardware events are cycles, instructions, branches, branch misses,
page faults, context switches and CPU migrations; their whole-child scope
cannot establish operation-local attribution. The primary admission rule
requires at least 5% p50 improvement in both total and commit for every
shape/repeat, plus exact semantic/output/provider/resource oracles. Every
over-5% adverse latency, mean, RSS or incremental-peak row is retained and
reviewed. The supplemental eager confirmation passed its median/mean guard
with zero >5% matched or same-build drift flags; it does not establish general
eager acceleration or host stability.

## Reproduction, when the source is frozen

The capture sequence is intentionally serial and must use the owned paths in
`plan.json` with Rust 1.95.0, two build jobs and nonincremental builds. The
source-bound preflight has passed and the canonical candidate manifest is
frozen. A replay must follow the frozen driver and preserve every receipt,
failure and raw report; this recipe does not assert the historical command
chronology:

```sh
python3 -B docs/performance/results/change-0525/preflight.py preflight-N
python3 -B docs/performance/results/change-0525/run.py candidate freeze
python3 -B docs/performance/results/change-0525/run.py candidate build-normal
python3 -B docs/performance/results/change-0525/run.py candidate native-r1
python3 -B docs/performance/results/change-0525/run.py candidate native-r2
python3 -B docs/performance/results/change-0525/run.py baseline native-r2
python3 -B docs/performance/results/change-0525/run.py candidate build-alloc
python3 -B docs/performance/results/change-0525/run.py candidate alloc
python3 -B docs/performance/results/change-0525/run.py candidate profile
python3 -B docs/performance/results/change-0525/run.py candidate hardware
python3 -B docs/performance/results/change-0525/eager_guard.py candidate capture
python3 -B docs/performance/results/change-0525/remaining_checks.py
```

The baseline r1, baseline allocator/profile/hardware/eager lanes and their
source-bound artifacts already exist in `baseline/`. The retained baseline
native r2 control was executed after candidate native r1/r2 as required by
ABBA; its receipts and manifests are authoritative. Do not interpret the
command list as evidence that a candidate build or capture has completed;
receipts and manifests are authoritative. The current native, allocator,
profile, hardware, original eager and supplemental eager captures are complete.
The supplemental confirmation has four children and 400 samples; its receipts
and comparison are authoritative.

## Evidence custody and status

The baseline source inventory has 451 XLSX files and manifest SHA-256
`4838d157514eb9479d080d475f32921f97c30acc3ca59da43ceb1626cd9c5d03`; its
source patch is empty. The frozen plan SHA-256 is
`78d3fad228e4bd00476148044bbcffc8d48ec02f1e825b59c95385f92da40f97`, and the
capture driver SHA-256 is
`1415937a8ee4602d5d8205c1b7fb98b84fd3d002d44758d3c5b8fe84a283b75d`. The
supplemental eager plan SHA-256 is
`08bd39c750ef27ef4a337b603f67e432f40f59c8c3a5e424a1d0758d388b0eb9`, and its
capture wrapper SHA-256 is
`2485c9f70cdb475c1eed0bc25d266de3a8ab18739090d1d6c53ae4aaa2e36ca4`. The
canonical candidate manifest is
`0fab9bafc238611761659bcafdd2e7b5df2543aecb194ac5477c168a2d6e6096`; its
exact candidate diff must stay within `crates/litchi-xlsx/`, while harness
source stays identical. Stage-local manifests, source patches, binary hashes,
commands, reports, logs, allocator identities and owned scratch paths are
part of the replay contract.

`verify.py` is expected to report `incomplete` until candidate and retained-A2
source bindings, both complete native stages, all lane analyses, the accepted
disposition, cleanup and the recursive evidence seal exist. Quality receipts
and the supplemental eager review now pass, while a failed or partial preflight
remains useful evidence but cannot be aliased to a quality pass. The eager
analyzer has a separate schema because eager reports intentionally have no
source phase; the failed generic analyzer attempt and its correction are
retained in `eager-guard-verification-adjustment.json`.

The first candidate preflight attempt stopped at compilation because a new
facade re-export was missing. After that fix, the second compiled and ran 978
unit tests, with two new tests failing during fixture construction because
formatting whitespace caused a `NotCompact` PackageWriter error. Root fixed
the raw ZIP fixture construction to preserve its indentation without changing
production behavior. The third attempt passed 1,288 test executions and XLSX
Clippy. These attempts do not constitute candidate measurement or acceptance.

`remaining_checks.py` proves byte-for-byte equality between the preflight-3 and
canonical candidate manifests and reuses the successful XLSX suite and Clippy
receipts. It runs the other ten frozen quality commands serially. All 12
quality gates pass, with 1,293 successful quality test executions. The
supplemental eager review and accepted disposition are recorded. Cleanup and
the strict post-cleanup verifier pass, with exact report replay and owned paths
absent. The recursive `SHA256SUMS` inventory covers 550 evidence files;
`verify.py --sealed --strict` passes against that inventory.

The separate supplemental validator probes also pass: a valid retained
100-sample control is accepted, while altered p50 and sink-vector probes are
rejected. These custody/schema probes are outside the 1,293 quality-execution
count.

No physical-provider, cold/range, native-producer, fuzz, scaling or broad CRUD
claim is made by this scoped result. The measured native, profile and
allocator values above are the candidate comparison; the original eager row
shift and all main adverse rows remain retained. The five global performance
ledgers carry the accepted 0525 decision. Cleanup, exact replay and recursive
seal verification are complete. The overall OLE2/OOXML goal remains open.
