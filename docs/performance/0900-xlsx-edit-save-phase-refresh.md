# 0900 — current XLSX one-percent edit/save phase baseline

Publication is the largest measured phase on both current source-backed XLSX
shapes: **43.88–44.33%** of the medium timed interval and **46.26–46.81%**
of dense-sparse. Planning follows at **33.89–34.40%** and **32.53–32.79%**.
All twelve native children / 1,200 measured samples pass semantic, preservation,
resource, and numerical replay. Two qualification children / eighty samples
remain separate. No production change or speedup is claimed; iWork is excluded.

## Why this measurement

[0899](0899-xml-validation-cost-attribution.md) did not establish sufficient
benefit for another small legacy-MCE rewrite. The broad goal requires ranking
larger CRUD work as well. [0705](0705-xlsx-edit-save-baseline-refresh.md)
measured this exact edit/save selector, but subsequent publication auditing,
XML validation/hardening, and runtime compression changes invalidate its old
cost ranking as current evidence. Later total edit/save measurements do not
supply this phase decomposition. Historical times are not a before leg.

The current executable is the retained 0894 baseline release build. All
10,771 selected source files match its source manifest by size and SHA-256;
selected tracked inputs are unchanged from its `506aa2edae` source revision
through evaluation HEAD `0d9df1ec3d`. The executable and original successful
build receipt are independently hash-bound. No new Rust build was necessary.

## Scope, corpus, and protocol

The existing selector is
`xlsx_source_backed_cell_values_one_percent_edit_save`. It uses fresh editors
and caches over instrumented in-memory positional sources, with all four
worksheets selected. Medium contains 9,216 cells and changes 93; dense-sparse
contains 17,792 and changes 178. Each package has seventeen members, including
eight deterministic 512-KiB media payloads. Twelve members remain untouched.
This is neither a producer-shaped MCE workload nor physical cold-cache evidence.

The total is the sum of four measured public phases: open, selector planning,
staged sets plus commit, and stream publication. Publication includes consuming
the editor and dropping the returned publication snapshot before the timer
stops. Sink preparation, remaining handles and sink destruction, source checks,
resource observations, output reopen, and semantic/preservation oracles are
outside that total. Reopen is recorded separately; it is substantial and must
not be silently counted as zero. The [source audit](results/change-0900/source-scope.md)
provides code references and the exact boundaries.

The frozen protocol has six blocks, alternating forward/reverse shape order.
Each child is a fresh process pinned to CPU 12, with ten warmups and one hundred
measured samples. Before those blocks, each shape has a separate two-warmup,
forty-sample qualification. Commands execute sequentially, stop on failure,
and never replace slow rows. The host is a shared AMD EPYC 9R45 under Linux;
affinity is not exclusive access. The retained build uses Rust 1.95, release
optimization level 3, fat LTO, one codegen unit, debug information, and the
system allocator. Exact source, build, host, environment, commands, and raw
samples are in the [packet](results/change-0900/README.md).

## Native results

Values below are medians of six child statistics. They are not pooled-sample
quantiles. Ranges describe observed children, not confidence intervals.

| Shape | Total p50 ms | p50 child range ms | Total p95 ms | Total p99 ms |
| --- | ---: | ---: | ---: | ---: |
| Medium | 16.655520 | 16.615693–16.704893 | 16.783212 | 16.862702 |
| Dense-sparse | 33.026613 | 32.998012–33.097997 | 33.341219 | 33.476989 |

The frozen child-p50 max/min review threshold is 1.20. Observed ratios are
1.005368 and 1.003030; neither flags. This does not establish universal tail
stability. Each report retains its Student-t mean interval, recomputed during
replay; no historical speedup interval or cross-host inference is made.

| Phase | Medium median child p50 ms | Dense-sparse median child p50 ms |
| --- | ---: | ---: |
| Open | 0.081725 | 0.100428 |
| Planning | 5.683822 | 10.770209 |
| Staging plus commit | 3.521477 | 6.772321 |
| Publication, including returned snapshot drop | 7.362020 | 15.414665 |
| Separate untimed reopen/verification diagnostic | 25.182893 | 29.065386 |

Phase quantiles are not additive. The reducer instead recomputes every sample's
four-phase sum and aligns it with the report's sorted elapsed vector and
original sample order. Per-child phase shares divide summed phase durations
by summed total durations. Commit shares are 20.92–21.42% / 20.15–20.64%; open
shares are 0.49–0.54% / 0.30–0.33%. Publication ranks first in every child.

## Preservation and resource evidence

Corpus digests are `dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036`
(medium) and `893ad3f5dd6a98aec44bc541a140048072c84c579b4b9e332431f779b097cb1a`
(dense-sparse). Output digests, semantic digests, untouched-member digests,
source counters, cache observations, and sink observations repeat across all
qualification and native samples for each shape.

| Observation per workflow | Medium | Dense-sparse |
| --- | ---: | ---: |
| Source read calls | 183 | 183 |
| Logical source bytes read | 4,233,170 | 4,258,604 |
| Payload materializations | 6 | 6 |
| Untouched members | 12 | 12 |
| Output bytes accepted | 4,226,788 | 4,252,436 |
| Sink write calls | 143 | 143 |
| Largest sink write bytes | 65,536 | 65,536 |

Logical read counters are not device I/O or decompression counts. Zero
unselected-sheet reads is not evidence of selective-subset scaling because
all worksheets were selected. This unmanaged control has unavailable native
allocation measurements and no active managed-budget refusal observation;
zero placeholder fields do not prove a resource-refusal test. No new memory,
RSS, copied-byte, compression, instruction, or parallel-scaling claim follows.

## Retained validator failure and recovery

The first qualification command succeeded and produced valid workload output.
The first Python validator then incorrectly required the report's Git revision
to equal the evaluation HEAD. The harness actually queries Git under its
compile-time `CARGO_MANIFEST_DIR`; the isolated 0894 source directory was already
cleaned. Git revision and dirty status are therefore unavailable. The fixed
child PATH also has no Rust compiler, so runtime `rustc_version` is null.
Source provenance comes from the manifest, build receipt, and executable hash,
not those optional runtime observations.

`run.py`, the original command, raw report, and `validation-failure.json` remain
unchanged. Additive `run_v2.py` explicitly checks the unavailable Git metadata,
revalidates the existing successful child, and runs only the remaining schedule.
The first child remains bound to its original helper; subsequent children bind
the second helper. No workload was repeated. The frozen plan is unchanged;
metadata and timing clarifications are additive. Written methodology review
followed qualification and overlapped capture; it is not represented as prior
written approval.

The retained 0705 validator is reused only for its source-independent semantic,
resource, phase-alignment, and numerical checks, with exact helper hash binding.
No historical timing enters this baseline. Final independent review, fresh
byte-identical analysis replay, and root closure bindings are retained in the
packet. No production code changed, so production test/build gates were not
rerun. No worktrees or build roots were created; the owned duplicate replay
output was removed, and unrelated workspace work remains untouched.

## Next measured boundary

The next target is current publication-owner profiling on these exact two
shapes, separating source-compatible XML audits, preservation writing,
compression, and other owner work without summing overlapping inclusive costs.
Planning remains the second target. Historical audit reuse was primarily a
one-cell benefit; scattered one-percent edits did not inherit that gain.
Neither required audit may be skipped merely because semantic readback passes.
No optimization is adopted from the phase ranking alone, and the broader
non-iWork performance goal remains open.
