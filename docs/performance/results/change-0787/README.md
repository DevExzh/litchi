# 0787: cached ordered Part scheduling

**Disposition: rejected.** Small/primed/floor-0/width-4 whole-child RSS fails
the frozen guard (+5.70%, CI above zero). Sixteen cached latency benefits do
not override that rule. Production is restored byte-for-byte to the baseline;
see `disposition.json` and `restored-source.json`.

This packet compares the baseline at `6b47617020` with a private cache-read
scheduling candidate. The public `tools/perf-execution` benchmark stays byte
identical to change 0786. The candidate and its focused tests are archived
under `candidate/`; build manifests identify the exact compiled sources.

## Reproduction and interpretation

`plan.json` fixes 60 cases: fresh/primed source-backed Parts, small/large/mixed
payloads, task floors 0/65536, and requested widths 1/2/4/8/32. The six native
blocks pair before/after children in fixed alternating orders. Each child has
30 samples and three warmups. Separate observer and qualification lanes
verify output, source reads and finite resource accounting. Observer timings
are never pooled with native results.

`adoption-policy.json` was frozen before builds. The pre-capture
`policy-interpretation.json` records its conjunctive regression rule: a case
rejects when the median paired after/before p50 or RSS ratio exceeds 1.05
and its bootstrap 95% lower endpoint exceeds 1. Benefit requires an eligible
primed case with at least 3% median improvement and an upper endpoint below 1.
All correctness gates remain mandatory; individual tails remain visible.

Root-only runners are `build.py`, `capture.py`, and `quality.py`. Captures
require the historical worktree and target paths in `origin.json`. Offline
replay resolves historical artifact paths against this packet after cleanup:

```sh
python3 -B docs/performance/results/change-0787/raw_audit.py --check
python3 -B docs/performance/results/change-0787/analyze.py --check
python3 -B docs/performance/results/change-0787/profile_analysis.py --check
python3 -B docs/performance/results/change-0787/validate.py --require-final-seal
```

The final seal checks both the exact payload file set and each SHA-256.
Executable hashes are checked before removal; `cleanup.json` retains that
custody witness. Raw logs and patches use packet-local Git attributes to
preserve bytes.

## Profiling is diagnostic

Baseline Callgrind uses a separate probe with the same corpus and an explicit
operation wrapper. The first attempt failed before the region in rustix CPU
clock initialization; its zero-count failure and original sources are retained.
The compatibility probe disables only that CPU clock. Native benchmark
clock handling is unchanged. Guest instruction counts are not native CPU
shares or a physical serial fraction; cross-thread collection is qualified
separately. Before/after syscall traces include priming and the measured read,
so their thread counts describe whole children, not isolated wall intervals.

The scope is a low-level in-memory source read control. It does not establish
whole Office CRUD speedups, physical cold behavior, delayed remote reads, or
cross-session contention.

Independent reviews: [source review](source-review.md) and [results review](results-review.md).
