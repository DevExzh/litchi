# ODS trigonometric and hyperbolic evaluation

This batch adds the 24 trigonometric, hyperbolic and angle/constant functions
listed in the [contract](contract.md) to the existing explicit OpenFormula
evaluator. Scalar expressions and elementwise array/local-reference expressions
use the same numerical kernels. No new dependency or public configuration type
is needed. Evaluation remains read-only; opening or saving a workbook never
recalculates formulas or publishes cached values.

The finite binary64 profile checks domains and unrepresentable results,
preserves formula-error propagation and lazy branches, and shares the existing
work, storage and cancellation budgets. Numerical kernels use fixed storage.
Reciprocal hyperbolic functions avoid intermediate overflow and premature
underflow; inverse cotangent follows OpenFormula's principal branch.

The [native observations](native/README.md) retain 115 selected LibreOffice
numeric caches across all 24 functions, with source hashes and a reproducible
extractor. They corroborate the authored vectors but do not establish native
application acceptance. The permitted `ATAN2(0;0)` profile variance is recorded
explicitly.

The [high-precision oracle](numeric_oracle.py) independently reproduces the
extreme-value [goldens](numeric-goldens.json) using mpmath 1.3.0 at 600 decimal
places with exact binary64 inputs. Selected transcendental vectors allow four
ULPs for platform rounding; signed zeros and the minimum-subnormal tail remain
exact checks. This evidence tool adds no library dependency.

Independent [review](review.md) passed. Final isolated validation passed all
1,193 ODS tests, with zero failures or ignored tests, including 20 new public
tests. Strict all-target Clippy, warnings-denied rustdoc, complete crate
formatting, changed-file formatting, dependency boundaries and diff checks all
passed. The [gate receipts](gates/results.json) and
[source verification](gates/verification.json) record the exact commands, logs,
lockfile and stable source manifests.

The paired [performance report](performance/results/performance-report.md)
retains 1,800 fresh-process observations: 300 baseline and 1,500 candidate,
with three warmups per child and 15 samples per workload/phase. All 24 new
functions have scalar workloads; five representative unary functions also have
4×4 and 16×16 array workloads, with additional synthetic local-reference cases.
Across the 20 matched arithmetic/ROUND control groups, median elapsed changes
range from −3.3% to +3.3%. Allocation calls, requested bytes and peak live-byte
deltas match exactly; median RSS changes range from −6.2% to +2.7%. The [final review note](performance/results/final-review.md) records one
plan variance: no separate baseline trigonometric-refusal receipt was captured.
These local
measurements do not establish a general speedup or full-workbook performance.
All measured result drops restore allocator live bytes. The independent
[verification receipt](root-verification.json) checks raw samples, published
medians, source identity, lockfiles and compiler settings.

This is a subset of §6.16 and does not complete OpenFormula, dependency
recalculation, or the overall spec-gap audit.

## Reproduction

The gate runner takes a source checkout and an external Cargo target directory.
Copy `gates/` to a scratch directory before running its `run.py` to retain the
committed receipts. Use the retained `gates/Cargo.lock` in the source checkout;
`batch-files.json` identifies the changed Rust files, and `freeze.json` identifies
the exact tested source and native-cache input. The runner checks all ODS tests,
strict Clippy and rustdoc, formatting, dependency boundaries and diff hygiene.
`verify.py` checks retained log hashes, unchanged source manifests, current
selected source hashes, test totals, the native observation catalogue and the
raw performance samples and reported statistics.

The separate performance directory retains the harness, profile, runner and
raw observations needed to reproduce its workload measurements. Its numerical
preflight checks evaluator projection against direct numerical reference
operations; the independent high-precision and upstream-cache vectors live in
the public integration tests.
