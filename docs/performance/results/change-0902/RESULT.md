# 0902 final evidence index

**Rejected.** Both primary cases failed the frozen end-to-end benefit gate.
Production and test files are restored to exact HEAD/baseline bytes; only the
experiment and documentation are committed. See the
[change report](../../0902-xlsx-upfront-stopped-audit-candidate.md),
[decision](decision.json), [canonical analysis](analysis.json), and
[descriptive tables](summary.json).

`README.md` is the preserved pre-capture helper protocol reviewed before
execution. Its v1 commands are historical. The final executed capture helper
is `run_v2.py`; the final acceptance analyzer is `analyze_v3.py`. `plan.json`
and `METHOD.md` retain the original schedule and gates, unchanged by the
versioned implementation corrections.

## Evidence and execution

- `source-final.json`, `candidate-source-v2.json`, `candidate-overlay/`, and
  `candidate.patch` bind the tested two-file candidate. The prior unformatted
  manifest and compile-failing patch remain separately preserved.
- `quality-summary.json` indexes every quality attempt. Final XML, Office,
  documentation, formatting, warning-denied lint/doc, and boundary gates pass.
- `candidate-*-build-run.json` and `candidate-*-binary.json` bind successful
  builds. Baseline descriptors point to retained 0894 source and binaries.
- `qualification/` contains 8 children / 16 samples; `native/` contains 48
  children / 4,800 samples; `allocator/` contains 16 children / 48 samples.
  Each child retains command, report, stdout/stderr, resource and receipt data.
- `qualification-analysis.json` passed before `native-capture-start.json`.
  `analysis.json` admits all three populations and rejects adoption. Native
  benefit statistics exclude qualification and allocator timings.
- `source-review.md`, `helper-review.md`, `helper-v2-review.md`,
  `helper-v3-review.md`, and `adr-review.md` record source and method review.
  `result-review.json` and `result-review.md` independently reproduce the raw
  native reduction and review allocator evidence (60 measured vectors / 180
  phase samples; four DOCX records have no phase metrics).
- `helper-v2-amendment.json`, `helper-v3-amendment.json`, failed command logs,
  `candidate-v1.patch`, and `summarize-v1-failure.json` preserve the failure
  history. No completed workload was rerun or replaced.

The prelaunch v1 attempt left only an empty directory and therefore no child
report. Its location is recorded in `helper-v2-amendment.json`; Git does not
store empty directories. Both the failure log and supervised command receipt
are retained.

## Replay after cleanup

The live-capture analyzer deliberately requires the candidate source in both
original roots. Those roots are no longer a prerequisite for historical replay:

```sh
python3 -B docs/performance/results/change-0902/replay.py \
  --output /tmp/litchi-0902-replay.json
```

Use an absent output path: helpers refuse overwrites. Replay verifies all
10,771 baseline archive file hashes, both retained candidate overrides, binary
and build bindings, the 72 raw child reports/receipts, semantic identities and
the exact canonical reduction. It does not claim the current checkout still
contains the rejected candidate. Baseline `change-0894/source.tar.gz`, baseline
binaries, and `change-0902/binaries/` are deliberately retained local binary
artifacts excluded from Git; their exact hashes and sizes are in the packet.
They are required for this full custody replay.

Both `raw-replay.json` and `post-cleanup-replay.json` pass. The latter ran after
source restoration and deletion of both temporary roots. `cleanup.json`
records removal of 7,314,983,405 logical bytes from the marked source and build
directories. Fixture links were unlinked without following them. This batch
created no Git worktrees and removed no unrelated worktrees. Existing unrelated
untracked files are bound by `unrelated-preservation.json`.

The experiment finds no qualifying primary benefit. It does not prove all
observer costs are zero, establish why the DOCX control moved, attribute the
RSS increase, or satisfy the broader Office performance goal.
