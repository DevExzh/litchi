# 0626: the CI smoke check compares against the last full run, and the comparison it was running had been failing closed for 273 commits

Status: retained, workflow and Python tooling only. No file under `crates/` or
`tools/perf-baseline/` changed. `performance_claim: none` — no claim-registry
entry; nothing here is a measurement of the library.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred
until that goal completes; iWork is excluded.

## What was changed

Evidence gap 1 of change 0587's survey: "the CI smoke check cannot detect a
regression". The `smoke` job of `.github/workflows/perf-baseline.yml` built its
comparison baseline by copying the report it had just produced and changing only
the recorded git revision, then asserted the comparator said `pass` against
itself. That is a working plumbing check and it cannot detect anything, because
both sides are the same measurement.

Four pieces:

1. **The `full` job publishes a reference.** It now also captures the bounded
   allocator report the smoke job captures — the same `litchi-perf-baseline-alloc`
   release binary, `--warmup 3 --samples 15 --case opc_file_eager_open
   --filesystem-cache warm,cold-requested` — as `target/perf/allocator-baseline.json`,
   writes `target/perf/allocator-baseline-descriptor.json` beside it, and ships
   both in the existing `container-performance-baseline-<run id>` artifact. The
   descriptor names the run, its event, its runner labels, its git revision, the
   harness binary profile and the case/corpus key manifest digest.

2. **The `smoke` job fetches it.** A new step lists recent successful runs of
   this workflow on `main` with `gh run list`, hands the listing to
   `perf_smoke_baseline.py choose-runs`, and downloads the newest candidate's
   artifact with `gh run download`. The step is `continue-on-error: true`, runs
   under `set +e`, and records what it achieved in a `fetch-status.json` written
   by the tool rather than by a shell heredoc, so the selector never has to
   interpret an exit code or a missing directory.

3. **A new module decides, classifies and reports.** `tools/perf_smoke_baseline.py`
   has five subcommands — `descriptor`, `fetch-status`, `choose-runs`, `select`
   and `report` — and holds every branch of the decision. `select` uses the fetched
   report only when it is bound to the same comparator policy identity, the same
   case/corpus key manifest digest, the same harness binary profile and the same
   runner labels, and when its revision differs from the commit under test and
   its worktree was clean. Otherwise it writes the old self-comparison baseline
   and labels it, in the job summary and in the annotation, as a plumbing check
   that detects no regression, listing every reason the fetched reference was
   rejected. `report` classifies the comparator's verdict and renders the
   annotation and the job summary.

4. **The comparison is advisory, and the expectations moved into a checked
   document.** `docs/performance/perf-smoke-baseline-policy-v1.json` carries the
   bindings, the self-comparison's expected shape (the seven assertions that used
   to live in a workflow heredoc) and `enforcement: advisory`. The smoke
   comparison artifacts upload as `container-performance-smoke-comparison-<run id>`.

## The defect this uncovered

The self-comparison the survey called "a real, working plumbing check" has not
run since commit `126c4a8b2` (change 0421's peak-observation fix, 273 commits
before this one). `perf_compare._validate_report_identity` compares the report's
complete `tool` object against `policy.tool_identity` for equality. Since
`126c4a8b2` the allocator harness emits
`tool.allocator_counter_revision` — `post_update_peak_v2` then, today
`serialized_region_peak_v3` from `tools/perf-baseline/src/allocation_metrics.rs`.
`docs/performance/perf-regression-policy-allocator-v1.json` was last edited at
`a878aceac`, two weeks earlier, and omits the field. So every freshly captured
allocator report is rejected with `baseline.tool does not match the policy tool
identity`, `perf_compare.py` exits 2, and the step fails.

Change 0421 anticipated the situation and left the policy markerless
deliberately: "Legacy markerless policies can still replay historical pairs for
schema compatibility... Both sides must be freshly captured with corrected
instrumentation before a policy opts into the new identity." In the smoke job
both sides are freshly captured by construction, so this change makes the
checked policy opt in: `tool_identity.allocator_counter_revision` is pinned to
`serialized_region_peak_v3`. That is the strict direction — a report from an
older allocator binary can no longer be compared under this policy — and
`tools/test_perf_smoke_baseline.py` asserts the pinned value against the string
in the harness source, so the two cannot drift apart silently again.

The other allocator policies are untouched.
`perf-regression-policy-xlsx-allocator-v1.json` and
`perf-regression-policy-opc-source-materialize-allocator-v1.json` are read by
retained evidence, not by CI; changing them would reinterpret that evidence.

## Why it is sound

**The comparison withholds every latency claim.** The smoke comparison runs
under the allocator policy, whose instrumentation makes `perf_compare` report
`latency_claims: withheld_instrumentation` and compare zero latency results. The
twenty compared metrics are deterministic allocation counters — calls, bytes,
live and peak before and after, over two cache states. `docs/GOAL.md`
deliverable 8's "do not make noisy cloud-hosted microbenchmarks a hard merge
gate until variance is understood" is therefore satisfied twice over: by the
metric class, and by `enforcement: advisory`, under which a regression or an
unusable comparison annotates the run and does not fail it.

**Infrastructure drift is named, not hidden.** A hosted runner may change CPU
model, kernel, memory size or toolchain between two runs, and the comparator
fails closed on each of those as a build-identity mismatch. A verdict whose
every error line begins `build identity mismatch for ` (or is the revision
collision) is classified `reference_environment_drift`, annotated as a notice
and never blocking, even under `blocking` enforcement. Any other error text is
an input defect. The classifier matches two literal prefixes from
`perf_compare`; an unrecognised message degrades to `input_defect`, which
annotates more loudly rather than less.

**The plumbing check keeps its teeth.** A broken self-comparison always fails
the job, whatever the enforcement setting, because a comparator that cannot
compare a report with itself is a tooling defect and not a measurement. The
seven assertions the old heredoc made are unchanged in value; they now live in
the policy document and are applied by tested code.

**A forged or mismatched reference cannot be used.** The case/corpus key
manifest digest is recomputed from the fetched report with
`perf_compare.report_result_key_manifest_sha256` and compared to the policy's
`expected_result_keys_sha256`; the descriptor's copy must agree with the
recomputation, and the descriptor's run id must equal the run the artifact was
downloaded from. Nothing in the fetched artifact is trusted on its own word.

**Nothing in the library moved.** `git diff 7082a1a3f -- crates/
tools/perf-baseline/` is empty. No ADR is engaged: no contract, limit, refusal
or output byte is touched by a CI workflow.

## Measured

Nothing here is a performance measurement. The numbers below are deterministic
counts and exit statuses.

Two real reports from the release allocator harness, each captured on a clean
worktree with `taskset -c 21`, at two distinct revisions:
`7082a1a3f` (the base commit, standing in for the last successful `full` run)
and `77220b62b` (this change's implementation, standing in for the commit under
test). `docs/performance/results/change-0626/verify-pipeline.py` drives the
descriptor, fetch-status, select, `perf_compare` and report steps over them,
once per scenario:

| scenario | selected mode | comparator | comparator exit | classified outcome | job exit |
| --- | --- | --- | ---: | --- | ---: |
| no artifact fetched | `self_comparison` | pass | 0 | `plumbing_pass` | 0 |
| compatible reference | `fetched_reference` | pass | 0 | `reference_pass` | 0 |
| allocation counters +10% | `fetched_reference` | regression | 1 | `reference_regression` | 0 |
| reference CPU model changed | `fetched_reference` | invalid | 2 | `reference_environment_drift` | 0 |
| reference over another corpus | `self_comparison` | pass | 0 | `plumbing_pass` | 0 |

The third row is the point of the change: a 10% rise in every allocation counter
of the current report is detected against a real prior report and reported
without failing the job. The fifth row falls back before comparing, naming the
digest it computed (`36f44718…`) and the digest the policy pins
(`debb7009…`).

Over the two real legs the comparator reports `pass`, 2 matched results, 20
compared metrics, 0 regressions, `latency_claims: withheld_instrumentation`,
0 latency results compared and 2 excluded. Against the base commit's own copy of
the allocator policy the same pair is rejected outright:
`ComparisonInputError: baseline.tool does not match the policy tool identity`.

## Correctness evidence

`tools/test_perf_smoke_baseline.py` adds 121 tests across thirteen classes: 14
policy-validation tests, 4 on runner-label normalization, 4 on the fetch-status
contract, 9 on `choose-runs`, 5 on the self-comparison baseline, 7 on the
descriptor, 20 on reference compatibility, 9 on selection, 18 on classification,
5 on rendering, 15 command-line tests, 9 on the workflow wiring and 2 on the
pinned allocator policy.

Every branch of the new module is covered: the policy document's key, schema,
scalar, event, enforcement and expectation refusals; `choose-runs` over an empty
listing, a non-list, a non-object entry, unsuccessful runs, other events, the
commit under test, invalid run ids, ordering and the limit; the self-comparison
baseline's relabelling, its non-mutation of the input and its acceptance by the
comparator; descriptor construction, its three construction refusals and its
key and scalar validation; each of the thirteen compatibility checks failing in
isolation, and both policy switches that skip one; selection in both modes
including the multi-reason fallback and a real comparator regression; all six
classification outcomes under both enforcement settings; annotation levels,
newline and percent escaping, and the advisory note; and fifteen command-line
tests that run the five subcommands and `perf_compare.main` over real files in a
temporary directory, for the fallback path, the fetched path, a fetched
regression, an artifact without the report, an unreadable report, an unreadable
descriptor, an absent reference directory, a missing fetch status, a malformed
current report and a broken plumbing check.

`tools/test_perf_baseline_source_policy.py::test_allocator_target_is_executed_and_compared_by_ci`
followed the expectations to their new home: it still asserts the workflow runs
the allocator target with the same arguments, and now asserts the checked smoke
policy carries the comparison's expected shape.

Gates, with tails in `results/change-0626/gates.txt`: `cargo fmt --all --check`
clean; the workflow parses under PyYAML 6.0.3; `python3 -m unittest` over the
nine modules the smoke job runs — 466 tests — OK; `tools.test_perf_baseline_source_policy`
14 tests OK; `check_perf_claims.py --mode strict` 10 claims;
`check_report_claim_classification.py` 167 rows;
`validate_crud_coverage_index.py` 15 categories and 33 selectors;
`non_iwork_gate.py verify` 45 bulk tree roots and 35 facade safe trees.

One pre-existing failure, reproduced unchanged in a detached worktree at the
base commit with none of this change present:
`tools.test_perf_claims.ClaimRegistryStructuralTests.test_seed_registry_is_structurally_valid`
reports two evidence ids in the registry that its expected set omits,
`abba-0418-pptx-cross-copy-lifecycle` and
`abba-0467-xlsx-cell-attributes-fixed-checkout`. The other 48 tests in that
module pass and `check_perf_claims.py` itself exits 0. Nothing here touches the
registry.

`actionlint` is **not installed on this host**, so the workflow was not linted
by it. What stands in its place: PyYAML parses the document,
`tools/test_perf_workflow_policy.py` (unchanged and passing) enforces the
resource-safety policy over the new steps, and nine new tests in
`tools/test_perf_smoke_baseline.py` assert the wiring textually — that both
Cargo jobs declare `LITCHI_RUNNER_LABELS` equal to their own `runs-on`, that
both name the checked smoke policy, that the `full` job publishes the report and
the descriptor, that the smoke job fetches, selects, compares and reports, that
the fetch step carries `continue-on-error: true`, that the comparator step
captures its exit status instead of failing, that the new files are in both path
filters and that the new suite runs in the smoke job.

## Validation preserved

No validation, limit, refusal or defence is touched: the change is confined to
`.github/workflows/perf-baseline.yml`, two `tools/*.py` files, one new policy
document, one pinned field in an existing policy document and one documentation
file. The comparator's own fail-closed contract is unchanged — `perf_compare.py`
is not modified. What changed is what the workflow does with its verdict, and
only for the smoke job; the manual `reference-regression` job still fails on a
regression or an identity defect.

## Limitations

- **No GitHub Actions run was executed.** Neither `gh` nor `actionlint` is
  installed on this host, so the fetch step's shell — `gh run list`, `gh run
  download`, the candidate loop — has never run. Everything downstream of it was
  driven locally through the `fetch-status.json` contract it writes.
- **What the first CI run must show**: that `gh` is present on `ubuntu-latest`
  and that `${{ github.token }}` with `actions: read` can list runs and download
  another run's artifact, including from a fork pull request; that
  `container-performance-baseline-<run id>` is downloadable by name from a prior
  run; that `$GITHUB_STEP_SUMMARY` receives the rendered summary; and that the
  reference's build identity actually matches a later runner's often enough for
  the fetched mode to be useful rather than permanently drifting.
- **The first push after this lands will fall back**, correctly and visibly: no
  `full` run has ever published `allocator-baseline.json`, so no candidate
  artifact carries one. The fetched path begins working after the next scheduled
  or dispatched `full` run.
- **The reference is one case.** `opc_file_eager_open` over one pinned corpus,
  warm and cold-requested: twenty allocation counters. It is not the 201-result
  release matrix, which a pull-request job cannot afford, and it says nothing
  about any other scenario.
- **No latency, cycle, instruction, RSS, cold-cache, range-source, scaling or
  cross-platform result is claimed**, and no statement is made about whether the
  hosted runner's allocation counters are stable run to run — that is precisely
  what deliverable 8 says must be understood before enforcement could change.
- **The drift classifier is coupled to two literal message prefixes** in
  `perf_compare.py`. If those change, comparisons that are really infrastructure
  drift will be annotated as input defects until the prefixes are updated.
- **The `full` job pays one extra feature-enabled build** of the harness plus
  one 15-sample allocator run. Cargo keeps the feature and non-feature artifacts
  side by side, so the release matrix build is not invalidated, but the added
  wall time on a hosted runner was not measured.

## Retained evidence

[`results/change-0626/README.md`](results/change-0626/README.md).
