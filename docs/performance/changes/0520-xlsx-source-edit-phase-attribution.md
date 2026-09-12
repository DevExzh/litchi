# 0520: attribute source-backed XLSX one-percent edit/save

`performance_claim: none`

`claim_authorized: false`

This evidence-only batch measures the current source-backed scalar-cell
one-percent edit/save route on two deterministic workbook shapes. Production
and Rust harness source remain at `b776ccb2a55c7a2a22018d4498e78676c6648dde`.
It supplies current phase and commit-owner attribution, following the
[0519 priority review](../results/change-0519/next-priority-review.md).
It does not repeat or overturn the rejected 0282 eager/source comparison.

Four fresh CPU-2 children retain 100 samples after 20 warmups each. Every
sample passes the existing semantic, deterministic-output and untouched-member
checks; the four children also execute exact no-op, clear/remove,
foreign-lineage refusal, stale-patch refusal and inverse semantic restoration
gates outside the timer. Medium and dense-sparse do not execute the
vendor-extension-only partial-sink gate. Inverse restoration here is semantic
and package-identity evidence, not an asserted byte-exact inverse artifact.

## Native phase results

The reported interval sums open, selector planning, staged sets plus commit,
and sequential publication. Publication includes dropping the returned
snapshot; sink setup, remaining handle destruction, output reopen and oracles
are excluded. Source input is an instrumented in-memory `ReadAt` provider,
with a fresh editor/cache each iteration and warmed process/code state.
Default range and filesystem options printed in the report do not activate
those providers for this selector.

| Shape / child | Total p50 ms | p95 ms | p99 ms | Mean ms |
| --- | ---: | ---: | ---: | ---: |
| Medium / 1 | 27.965 | 28.244 | 28.296 | 27.975 |
| Medium / 2 | 27.438 | 27.731 | 27.946 | 27.444 |
| Dense-sparse / 1 | 53.699 | 54.372 | 54.762 | 53.708 |
| Dense-sparse / 2 | 52.799 | 53.716 | 53.933 | 52.849 |

Across these four children, commit accounts for 44.80–45.80% of aggregate
measured time, planning 27.94–28.97%, publication 24.89–26.79%, and open
0.19–0.35%. These are sums of phase times divided by summed total time,
not sums of unrelated phase percentiles. The raw acquisition-order vectors
reconcile exactly with the sorted total vector through `sample_order`.

All total-time repeat differences stay below the 5% review threshold, but
12 phase/metric comparisons exceed 5% in either direction. Medium publication
p50 differs by -7.41%; medium open p99 differs by +24.41%. The complete list
is retained in [analysis.json](../results/change-0520/analysis.json).
These same-build variations are not optimization regressions or improvements.
The two children per shape provide limited process replication; raw
within-child mean confidence intervals do not establish cross-host confidence.

## Isolated commit instructions

Four further fresh children run under Callgrind, one retained iteration and
no warmup each. Collection toggles only inside `MultiSourceEdit::commit`.
Counters reset on entry and dump on return. All dumps are retained: three
untimed lifecycle commits precede the fourth, timed one-percent commit.
Each selected fourth dump proves one positive direct edge from the benchmark
runner, with raw summary = incoming inclusive instructions = self + direct
callees. The profiling boundary excludes the staged `set` loop that is part
of native `commit_ns`.

| Shape / repeat | Commit Ir | Candidate reconstruction Ir | Rewrite Ir |
| --- | ---: | ---: | ---: |
| Medium / 1 | 220,444,234 | 144,706,503 | 75,165,697 |
| Medium / 2 | 219,884,708 | 144,117,632 | 75,193,767 |
| Dense-sparse / 1 | 420,889,873 | 273,459,157 | 146,612,230 |
| Dense-sparse / 2 | 420,886,196 | 273,460,877 | 146,615,866 |

`Snapshot::from_rewritten_source` is 64.97–65.64% of commit instructions;
`raw::worksheet::edit::package::rewrite` is 34.10–34.84%. The former performs
candidate XML validation and semantic parsing; the latter is dominated by
the source layout scan. Nested parsing/validation costs overlap these parents
and must not be added again. These instruction shares are not native time
shares or removable-work estimates. Valgrind's retained `brk segment overflow`
notices limit allocator interpretation; every profile child nevertheless
completed successfully with output, source/cache counters and corpus matching
its native shape. No allocator claim follows from these profiles.

## Resource and correctness scope

The corpora contain 9,216/17,792 cells, 93/178 selected updates, four touched
worksheets, 17 members and eight untouched 512 KiB media members. Corpus hashes
match the historical deterministic generators. All samples retain six cache
materializations with no failures, evictions or bypasses, deterministic logical
read counters, zero unmanaged budget counters, 12 untouched members and a
maximum 64 KiB accepted sink write. The read counters include publication
transfer and do not measure device I/O, decompression or physical copies.
All worksheets are selected by this one-percent update set, so zero unselected
worksheet reads is not evidence for a selective-subset workload.

No production, public API, dependency, unsafe-code, budget, validation or
preservation behavior changes. All 30 previously read ADR/index hashes remain
unchanged. This batch validates the release build and executable harness
oracles, plus raw evidence/accounting and exact report replay. It does
not claim a fresh full workspace test, fuzz, native Office, allocator/RSS,
cold/range provider, or parallel-scaling pass.

## Next action

Target the named candidate-validation/semantic-store and source-layout owners
for operation-local allocation attribution and a concrete work-elimination
design. Preserve exact source, limits, namespace/MCE behavior, error ordering,
candidate validation and semantic readback. The rejected 0514 source-pass and
0516 output-fusion designs remain rejected; these profiles alone do not
authorize reviving them. Planning and publication remain material alternative
targets if no safe and useful commit change survives its own before/after
measurement. No full-route speedup or coverage-catalog promotion is authorized.

OLE2/OOXML optimization remains active. ODF is deferred until that goal is
complete, and iWork remains excluded. Reproduction commands, raw artifacts,
and verifiers are in [the evidence bundle](../results/change-0520/README.md).
