# Evidence: change 0626, the CI smoke check's baseline

Change record:
[`0626-perf-ci-smoke-baseline-fetch.md`](../../0626-perf-ci-smoke-baseline-fetch.md).

Disposition: retained, workflow and Python tooling only. `performance_claim:
none`. **No file under `crates/` or `tools/perf-baseline/` was modified.**
Nothing in this packet is a timing measurement: every number is a deterministic
allocation counter, a digest, or a process exit status.

## Contents

| Path | What it is |
| --- | --- |
| `reports/full-run-allocator-baseline.json` | A real release `litchi-perf-baseline-alloc` report captured at the base commit `7082a1a3f` on a clean detached worktree, standing in for the artifact a successful `full` run publishes. |
| `reports/smoke-current.json` | The same harness at `77220b62b`, this change's implementation commit, on a clean worktree, standing in for the report the `smoke` job produces for the commit under test. |
| `policy-drift.txt` | The checked allocator policy applied to a freshly captured allocator report, at the base commit and after this change's one-field pin, plus the comparison of the two real legs. This is the reproduction of the pre-existing CI defect and of its fix. |
| `verify-pipeline.py` | The driver that plays the `full` job's descriptor step and the `smoke` job's select, compare and report steps over the two reports, once per scenario. Retained and rerunnable. |
| `pipeline/summary.json` | One row per scenario: selected mode, fallback reasons, comparator status and exit, classified outcome, whether it blocks, and the job's exit status. |
| `pipeline/<scenario>/transcript.txt` | Every command the driver ran for that scenario, its stdout and its exit status. |
| `pipeline/<scenario>/{fetch-status,selection,comparison,classification}.json` | The machine-readable artifacts the workflow writes and uploads. |
| `pipeline/<scenario>/{comparison.txt,outcome.md,step-summary.md}` | The comparator's human summary, the rendered outcome and the job summary the workflow appends to `$GITHUB_STEP_SUMMARY`. |
| `pipeline/<scenario>/reference/allocator-baseline-descriptor.json` | The descriptor the `full` job would upload beside its report. |
| `gates.txt` | The tail and exit status of each gate, then the one pre-existing failure reproduced on the untouched base commit. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `log-sections.md` | The four log paragraphs for the coordinator to merge. |

The scenarios are `fallback` (no artifact fetched), `fetched` (a compatible
reference), `regression` (the current report's allocation counters raised 10%),
`drift` (the reference runner's CPU model changed) and `incompatible` (the
reference captured over another corpus).

Each scenario's own copies of the two 117 KB reports and of the baseline the
selector chose were written to a scratch `--work` directory and are not
retained: they are byte-for-byte reproducible from `reports/` by rerunning the
driver.

## Provenance

Base commit `7082a1a3f480589c0025aad8925cae315a4cfd4b`
(`feat/office-format-completeness`); branch
`perf/0626-perf-ci-smoke-baseline-fetch`. Host: AMD EPYC 9R45, 32 cores,
123 GiB, Linux 7.0.0-1012-aws; rustc 1.95.0 (59807616e 2026-04-14); Python
3.14.4 with PyYAML 6.0.3.

Both harness legs were built `--release --locked --features allocator-metrics
--bin litchi-perf-baseline-alloc` in separate worktrees with their own external
`CARGO_TARGET_DIR`, and run pinned to CPU 21:

| leg | worktree | revision | binary sha256 |
| --- | --- | --- | --- |
| reference | `litchi-worktrees/0626-ref` (detached) | `7082a1a3f4805890…` | `79d4b7e1a2a6f1b7ee02b408d6f6d91106b2542a647be2a3dc972eec0a4a08f4` |
| current | `litchi-worktrees/0626` | `77220b62b9c14900…` | `10e2f24db7218f3097ed5151d17fed1f6ad8183851c067f778e5b29b081dc591` |

Report digests: `full-run-allocator-baseline.json`
`11b9232ceaada421af590277cc7522e261ee4f16d658c9d0684e150a5d77fb4a`,
`smoke-current.json`
`d05ea9d3a1b68c810425b9bb2afe87899156e532ce1d131e0625b5aaf8394e64`.

The two binaries differ because Cargo embeds the workspace path; the comparator
deliberately permits differing `binary_identity` descriptors between a baseline
and a candidate, and the smoke policy binds the build *profile*, not the digest.

`77220b62b` is the implementation commit this branch carried while the legs were
captured. The final commit on the branch amends it to add this packet and the
change record, so the committed hash differs from the revision recorded inside
`reports/smoke-current.json`; the tree is otherwise identical.

The host carried other agents throughout. That affects wall time only: no
timing is reported here, and each leg's `logical_cpus_available` is 1 because
both were pinned with `taskset -c 21`, which is why each report's
`execution_workers` is `[1]` on both sides.

## Reproducing

```sh
# the two legs (each from a clean worktree at its own revision)
CARGO_TARGET_DIR=<external> taskset -c 21 cargo run --release --locked \
  --manifest-path tools/perf-baseline/Cargo.toml \
  --features allocator-metrics --bin litchi-perf-baseline-alloc -- \
  --warmup 3 --samples 15 --case opc_file_eager_open \
  --filesystem-cache warm,cold-requested --json <out>.json

# the pipeline
python3 docs/performance/results/change-0626/verify-pipeline.py \
  --reference docs/performance/results/change-0626/reports/full-run-allocator-baseline.json \
  --current   docs/performance/results/change-0626/reports/smoke-current.json \
  --work      <scratch> \
  --out       docs/performance/results/change-0626/pipeline
```

The driver prints one line per scenario and writes `summary.json`. Rerunning it
against the retained reports reproduces every file under `pipeline/` except the
absolute paths inside each `transcript.txt`.

## What this packet does not establish

- **That the workflow runs.** Neither `gh` nor `actionlint` is installed on this
  host, so the fetch step's shell has never executed and the workflow has never
  been linted by `actionlint`. Everything downstream of the fetch was driven
  through the `fetch-status.json` contract that step writes.
- That a GitHub-hosted runner's build identity is stable enough between two runs
  for the fetched comparison to fire in practice. The `drift` scenario shows what
  happens when it is not; how often that happens is unknown until the first runs.
- Any latency, cycle, instruction, RSS, cold-cache, range-source, scaling or
  cross-platform result. The allocator policy withholds latency by construction.
- That `opc_file_eager_open` is representative of anything else. It is one case
  over one pinned corpus in two cache states: twenty deterministic counters.
- That the pre-existing `tools.test_perf_claims` registry failure recorded in
  `gates.txt` is unrelated to *any* change — only that it reproduces unchanged
  at the base commit with none of this change present.
