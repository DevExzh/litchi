# ODS elementary real mathematical evaluation

This batch adds `ABS`, `EXP`, `LN`, `LOG`, `LOG10`, `MOD`, `POWER`,
`QUOTIENT`, `SIGN`, `SQRT`, and `SQRTPI` to the explicit scalar and
array/local-reference evaluator. The [contract](contract.md) records the
normative sections and finite binary64 choices. `POWER` shares its numerical
kernel with the infix power operator. Kernels use fixed storage and the existing
budgeted evaluation interfaces; no dependency or public configuration type is
added. Evaluation is read-only and does not publish cached worksheet values.

`MOD` computes the remainder of the represented operands without forming a
potentially overflowing quotient or subtracting large rounded products.
`SQRTPI` factors the square root to handle large and subnormal finite inputs.
Domain errors remain formula values, including per-cell errors inside arrays;
cancellation and resource failures return no partial result.

The [native observations](native/README.md) retain 56 selected LibreOffice
numeric caches across all eleven functions. One large-quotient `MOD` cache is
explicitly excluded as producer variance. The [independent numerical oracle](numeric_oracle.py)
reproduces fifteen extreme [goldens](numeric-goldens.json) using exact-rational
remainders and mpmath 1.3.0 at 600 decimal places from exact binary64 inputs.
These tools add no library dependency. Transcendental comparisons allow four
ULPs for platform rounding; exact-rational vectors test the rounded result.

Independent [source and harness review](review.md) passed.
The full isolated ODS suite passed 1,211 tests, with zero failures or ignored
tests, including eighteen new public integration tests. Strict all-target
Clippy, warnings-denied rustdoc and both crate-wide and changed-file formatting
checks, dependency boundaries and diff hygiene passed. The [gate receipts](gates/results.json) retain exact commands and
logs against the selected source freeze and retained lockfile.

The paired [performance report](performance/results/performance-report.md)
retains 1,710 fresh-process observations: 450 baseline and 1,260 candidate,
with three warmups per child and fifteen samples per case/phase. Candidate
workloads cover all eleven scalar functions, five representative unary kernels
at 4×4 and 16×16, and synthetic local-reference cases. The thirty matched
arithmetic/ROUND/SIN control groups show median elapsed changes from −4.1% to
+2.5%, with identical allocation calls, requested bytes and peak live-byte
deltas. Median process RSS changes range from −2.5% to +9.1%; this is a
whole-process measurement and no specific cause is established for that spread.
All result drops restore allocator live bytes. These measurements do not
establish a general speedup or full-workbook performance. The independent
[verification receipt](root-verification.json) checks raw observations,
reported medians, source closure, locks, compiler settings and artifact hashes.

This is a subset of §6.16. Remaining mathematical and other OpenFormula
families, dependency recalculation, result spilling and cache publication
remain outside this batch.

## Reproduction

The gate runner accepts a source checkout and an external Cargo target.
Reconstruct the selected source identified by `gates/freeze.json` and use the
retained `gates/Cargo.lock`; copy the gate scripts to scratch before rerunning
so the committed receipts are preserved. The separate performance directory
retains its matched workload, harness, locks, raw observations and source
identity. `verify.py` independently checks retained gate logs, native counts,
source hashes, performance samples and published statistics.
