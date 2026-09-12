# 0527 XLSX row-primary-arena pilot — rejected

The candidate missed the frozen total-time gate in dense-sparse repeat 2.
Production is restored; four baseline-compatible regression tests remain.
The retained 0525 optimization is unchanged. OLE2/OOXML remains the active
priority, ODF is deferred, and iWork is excluded.

## Result and scope

See [the change record](../../changes/0527-xlsx-row-primary-arena-pilot.md)
and [comparison.json](comparison.json) for every matched result. The native
campaign contains 2,440 measured durations in 36 fresh children; the allocation
lane contains 40 samples in eight separate children. CPU 2, release Rust
1.95.0, two serially controlled Cargo build jobs, no incremental compilation,
200 primary samples with 20 warmups, and 30 guard samples with 10 warmups are
fixed in [plan.json](plan.json).

Total medians improved 1.8214–3.5854% and commit medians 5.7713–8.0431% in the
rejected candidate. Allocation calls fell 18.9078–19.8814%. Dense repeat 2
missed both the 2% total median and 2% total mean gates. The allocation and
commit gains do not override those failures. Profile, hardware and eager
lanes are explicitly unmeasured because the pilot failed. No production
speedup or instruction-count improvement is claimed by this batch.

All 47 matched adverse flags and 71 same-build drift flags above 5% are kept
in the raw comparison and individually reviewed. Small incremental allocation
peak increases remain visible despite the allocation-call reduction.

## Evidence and source custody

- [Frozen inputs](frozen-inputs.json), [plan](plan.json), [driver](run.py),
  [start state](start-state.json) and [ADR manifest](adr-manifest.json).
- `baseline/`: accepted source manifest, empty patch, normal/allocator build
  receipts and binary identities, both native repeats and allocator repeats.
- `candidate/`: exact six-file patch and source manifest, release/allocator
  builds and all candidate captures. Three runtime files exactly match the
  frozen [0526 draft](../change-0526/row-primary-arena.patch); three files add
  tests only.
- `preflight-1/`: successful candidate all-features test receipt and exact
  source snapshot, 1,294 successful test executions and zero ignored.
- `final/`: restored runtime source with four compatible regression tests,
  exact patch/manifest and all 12 independently executed quality checks,
  with 1,297 successful test executions.
- [Source review](source-review.md), [test handoff review](test-review.md),
  [test patch](tests.patch) and [candidate-only tests](arena-only-tests.patch).
- [Analyzer](analyze.py), [numeric review](numeric-review.md),
  [individual adverse review](adverse-review.json), [quality summary](quality-summary.json),
  [decision](decision.json) and [independent verifier](verify.py).

The baseline normal build succeeded, but the first binary copy hit the `/tmp`
user quota before any measurement. [storage-recovery.json](storage-recovery.json)
binds the original successful build and recovered hash-checked executable.
The exact owned scratch path became a symlink to disk-backed
`/home/zhuhe/litchi-goal-0527-target/retained-binaries`; the frozen driver and
capture paths did not change. The retained baseline binary was used for A2
under the candidate source checkout. Compiled-source and working-source
manifests are separately recorded in receipts. No failed timing sample was
replaced. Point process observations are retained; they do not prove continuous
machine isolation.

## Replay and cleanup

From the repository root:

```sh
python3 -B docs/performance/results/change-0527/analyze.py --stage compare --output /path/outside/bundle/comparison.json
python3 -B docs/performance/results/change-0527/verify.py --component precleanup --strict
python3 -B docs/performance/results/change-0527/verify.py --component all --sealed --strict
```

The final verifier checks exact source-patch replay, ADRs, build/capture
custody, raw numerical reproduction, every flag review, final quality,
disposition, owned-path cleanup and the exact SHA256SUMS inventory. Temporary
executables and targets are removed after pre-cleanup verification; their
hashes and receipts remain. The test agent also removed its minimal temporary
copies. Building and capturing anew requires a fresh output directory and
source freeze; the retained receipts must never be overwritten.

The overall optimization goal remains active. Quantify another avoidable
scanner cost before the next independent OOXML candidate; do not lower this
pilot's gates or reclassify the rejected layout change as retained.
