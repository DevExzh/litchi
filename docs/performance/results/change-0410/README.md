# Change 0410 evidence

`latency/` is the strict selected-cell ABBA package. `capture/` contains all
initial selected, edit/save and allocation reports and catalog sidecars.
`guard-recheck/` retains the diagnostic repeat prompted by the original eager
edit/save regression; neither edit/save block authorizes a latency claim.
`profile/` is a separate candidate residual CPU capture. `checks/` retains
build/test/validation commands and logs, including failed initial checks.

From the repository root, verify the published inventory and source-bound
capture without changing the bundle:

```sh
python3 docs/performance/results/change-0410/verify-artifacts.py
python3 docs/performance/results/change-0410/verify-capture.py \
  --root docs/performance/results/change-0410/capture \
  --output /tmp/litchi-0410-recomputed-verification.json
python3 tools/check_perf_claims.py \
  --registry docs/performance/claim-registry-v1.json \
  --repo-root . --evidence-root . --mode strict
```

Large originals are losslessly compressed with zstd. The manifest binds both
compressed files and original content hashes/sizes. `verify-capture.py` accepts
the `.json.zst` reports directly. To recompute a summary, decompress the four
relevant reports into a disposable directory and pass them to
`tools/perf_abba_summary.py` in A1/B1/B2/A2 order. Do not mix the initial and
diagnostic guard blocks or replace an adverse report with a repeat.

To repeat the experiment, create an isolated clean worktree, build the two
normal/allocator executable pairs using `checks/commands.txt`, and copy each
pair before rebuilding the other revision. Adapt the absolute worktree,
binary and output paths in `capture/run-leg.py`, keeping all protocol values
fixed. Invoke it separately as `a1`, `b1`, `b2`, `a2`. Each role's
`*-build-identity.json` records the original executable hashes and build IDs;
the binaries themselves are not retained. Independent rebuilds may have
different debug paths/build IDs and must record their own identities.

The exact original profiling command is in `profile/capture-commands.json`.
The final `profile/attribute_0410.py` reads the immutable command snapshot and
the recorded candidate Git blobs. For offline attribution, decompress the
profile inputs into a disposable directory, then run that script with
`--capture`, `--script`, `--report`, `--repo` and `--output` pointing to the
extracted inputs and a repository retaining the candidate commit. Keep the
final parser rather than `attribute_0410-initial.py`: the initial adaptation
was rejected for a mistyped corpus hash. `capture-profile.py` retains the
original capture orchestration, including its failed first postprocessing
attempt. The final parser also corrects inherited prose and records the MCE
source blob. Raw samples were not recaptured to fix postprocessing.

The raw perf data includes surrounding verification and child startup. The
SVG is a whole-process flamegraph; only the exact selected-cell ancestor
filter establishes the narrower sampled subset. Neither the SVG nor the
weighted periods constitute an operation-only timer or paired CPU claim.
