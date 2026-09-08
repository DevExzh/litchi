# 0467: eager XLSX cell attribute scan

This bundle compares control `33243bc2e6bd63ab056eb45493b0717818501f4b`
with candidate `87733cf3b86c5500ee7aee4cf6a4cb0f3cabf6a7`.
The candidate replaces five checked attribute scans per eager worksheet cell
with one checked scan and four fixed borrowed slots. See the
[source review](source-review.md) and
[per-change record](../../changes/0467-xlsx-cell-attributes.md).

Both executables were built inside clean, detached worktrees with Rust 1.98.1,
release debug level 1, frame pointers and unwind tables. Normal captures use
no external profiler. Build receipts, executable hashes, the matched source
inventory and compile-time fixture hashes authenticate the comparison.
`source-bindings.json` explicitly lists the sparse checkout omissions; the
harness uses its tracked `tools/perf-baseline/Cargo.lock`.

| Artifacts | Scope |
|---|---|
| `A1-clean`, `B1-clean`, `B2-clean`, `A2-clean` | Normal ABBA: six rows, 100 samples and five warmups per row |
| `A-full-clean`, `B-full-clean` | Full default guard: 201 rows, 15 samples and three warmups |
| `A-heap-clean`, `B-heap-clean` | Whole-process Heaptrack: dense one-percent update, five samples and one warmup |
| `A1-guard-clean`, `B1-guard-clean`, `B2-guard-clean`, `A2-guard-clean` | Supplemental 29-row guard: 100 samples and five warmups |
| `A1-fixed`, `B1-fixed`, `B2-fixed`, `A2-fixed` | Same-path qualification: dense XLSX primary and three DOC guards, 500 samples and five warmups |
| `A-full-fixed`, `B-full-fixed` | Same-path full default guard: 201 rows, 15 samples and three warmups |
| `protocol-500.json`, `capture-500.py` | Unexecuted separate-path proposal, superseded by `protocol-fixed.json` |
| `validation` | Rust tests, feature check, formatting, lint, rustdoc and boundary receipts |

Each process is pinned to CPU 2 with one worker. Task-owned heavy jobs are
serialized. Independent light source review and edits may overlap, and the
shared KVM host is not otherwise controlled. The timer includes ordinary
commit plus sequential write. Source open, edit staging, expected-output
construction, output reopening, semantic verification and teardown are outside
that timer. Heaptrack includes the whole process; instrumented time and RSS
are not normal-latency or operation-local memory measurements.

Initial `A1` and `A-full` captures are diagnostic only. They used a reused
binary whose compiled-in repository path caused the harness to report a dirty
worktree. `capture-correction.json`, the initial binding and initial driver are
retained. The initial descriptive comparison uses fresh builds in clean role worktrees.
Its repeated DOC payload-heavy guard regression motivated fresh control and
candidate builds from the same absolute `shared-tree` checkout path.
`build-path-review.json` retains that observation and its limits. The fixed-path
driver checks out each binary's exact revision before capture so the compiled
manifest path reports truthful clean Git metadata. Both build pairs and all
observations are retained.
The first clean control build lacked two sparse-checkout compile-time fixtures;
the failed build logs are retained as `control-build-incomplete.*`.

The validation log retains the whole-workspace formatting failure at one
unchanged, excluded iWork file. `validation/format-scope.json` authenticates
that scope; XLSX formatting passed. No iWork source was changed.

`fixed-qualification.json` retains the complete four-row ABBA, DOC guards,
process RSS and every fixed full-matrix comparison. `fixed-summary.json` is
its complete canonical ABBA summary. `fixed-package-inputs` contains explicit
primary-only report projections; the standard compressed `fixed-package`
contains the strict claim evidence. Original four-row reports remain intact.
`acceptance-review.json` records the retention decision, including unresolved
short full-matrix flags. No DOC or process-memory claim is registered.

The primary dense XLSX median falls 8.68% / 7.86% in the two pairings. Mean,
p95 and p99 also pass the existing 500-sample/drift policy. See the per-change
record for all lane medians, guard limits and allocation scope.

The bundle verifier passed before cleanup and in a fresh copy with only the
two standard comparison modules and regression policy, without Git or live
benchmark binaries. The verifier tests pass 11 cases; the analysis and fixed
qualification tests pass 7 and 9 cases. The strict registry checker accepts all
10 registered claims, including the four scoped 0467 statistics.
`cleanup-builds.json` records removal of all three clean worktrees and four
binary copies totaling 1,854,649,416 bytes. `SHA256SUMS` seals every regular
bundle file except the seal itself.

For replay from this checkout after finalization:

```sh
python3 -B docs/performance/results/change-0467/verify.py
python3 -B docs/performance/results/change-0467/test_analyze.py
python3 -B docs/performance/results/change-0467/test_verify.py
python3 -B docs/performance/results/change-0467/test_fixed_qualify.py
python3 -B tools/check_perf_claims.py --repo-root . --evidence-root . --mode strict
```

To recapture, create detached worktrees at the two recorded revisions under
`/tmp/litchi-goal-0467/{control,candidate}-tree`, using the same sparse paths:
`crates`, `tools`, `docs/adr`, `.github`, `.cargo`,
`test-data/poi/test-data/spreadsheet`, and `test-data/rtf`.
Copy the build/capture helpers and protocols into a fresh directory at the
same repository depth as this bundle. Build each role with `build.py ROLE`,
then run `capture.py LANE` in the frozen normal order, with the full and heap
lanes serialized separately. Helpers refuse existing lane directories and
build logs.
The separate-path 500-sample proposal was never executed.
For the authoritative qualification, create an additional clean detached
`/tmp/litchi-goal-0467/shared-tree` with the same sparse paths, build both roles
with `build-fixed.py ROLE`, and run `capture-fixed.py LANE` for the frozen
A1-fixed/B1-fixed/B2-fixed/A2-fixed order and both full-fixed guard lanes.
Use an external `flock /tmp/litchi-goal-0467/cpu.lock` around each build and the
complete capture pipeline. Run `fixed_qualify.py` to recompute the individual
primary and DOC guard outcomes. The full supplemental CFB matrix retains a
source-counter mismatch; `guard-summary.json` preserves the canonical rejection
and limits strict eligibility to unaffected projections.

Run `postprocess.py` only after captures finish and before deleting the
matching temporary binaries needed for symbolization.

The current default case matrix and checked CRUD coverage are unchanged.
This owned-byte, single-worker synthetic comparison does not establish
physical-cold, remote-source, bounded-streaming, native-application or scaling
results. The broader non-iWork performance goal remains open.
