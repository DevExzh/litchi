# ODS numeric aggregate validation

This batch adds `SUM`, `PRODUCT`, `SUMSQ`, `SUMPRODUCT`, `SUMX2MY2`,
`SUMX2PY2`, and `SUMXMY2` to the bounded read-only formula evaluators.
The resolver-free evaluator accepts scalar inputs; the value evaluator also
accepts supported arrays and references. Formula caches remain inert.

The [contract](contract.md) records the OpenFormula signatures, argument and
element conversion policies, formula-error ordering, and numeric profile.
NumberSequence references omit Text, Logical, and Empty cells; inline arrays
and forced matrices use the documented element conversion policy instead.
Matrix reducers require identical dimensions, with scalars treated as 1×1
matrices. `PRODUCT()` is an arity error; an explicitly supplied reference
containing no selected numbers has multiplicative identity one.

The numerical evidence has two independently generated sets:

- [17 targeted observations](numeric-goldens.json) cover cancellation,
  intermediate overflow and underflow, tiny square sums, and variadic product
  cancellation. [Generator](numeric_oracle.py).
- [336 seeded observations](differential-goldens.json) cover all seven
  functions over binary64 operands spanning subnormal through near-maximum
  exponents. [Generator](differential_oracle.py).

Both generators use Python's standard-library `fractions.Fraction` over the
exact represented operands. The Rust oracle test consumes these retained
observations. Tests permit a documented tolerance where normalized product
multiplication or subtraction rounds before the final sum; exact sums and
pair-product sums use bit comparisons. This is targeted numerical validation,
not a proof of every floating-point input.

[Native evidence](native/README.md) retains selected LibreOffice fixture
formulas, cached results, and bounded literal-cell closures at a pinned upstream
commit. No formula-cell cache is used as an input. These are comparisons with
recorded producer caches, not a native producer execution or save/reopen test.

The frozen isolated checkout passes all seven gates in [the retained
receipts](gates/results.json): 1,238 ODS tests (zero failures or ignored tests),
strict all-target Clippy, warnings-denied rustdoc, complete crate formatting,
selected-file formatting, dependency boundaries, and diff checks. The
[freeze](gates/freeze.json) includes the deletion of the old database numeric
module and the new shared module. Existing database accumulator state and
kernel bodies remain unchanged by that move. The old capability-refusal test
now uses the still-unimplemented `SUMIF` instead of the newly supported `SUM`.

The three-or-more-factor product-sum profile has a fixed 8,192-bit accumulator
window with 64 carry bits reserved. Exceeding the 8,128-bit admitted input
window is a typed memory-limit failure, including through `IFERROR`; it is
not a formula numeric error. Its receipt uses two-magnitude input bytes
(limit 2,032), independently of additional caller budgets. Intermediate
products use normalized binary64 mantissas and may round at each factor;
the retained normalized terms are then accumulated exactly.

The [performance report](performance/results/performance-report.md) retains
2,430 fresh-process samples: 360 baseline and 2,070 candidate observations,
with three warmups and fifteen samples per group. The 24 matched control
groups have identical allocator-call, requested-byte, and peak-live-byte
medians. Raw elapsed-time medians change by −1.46% to +1.73%; process RSS
medians change by −2.48% to +9.35%. New aggregate workloads are candidate-only
measurements because the baseline does not implement these functions.
Nested reference cases retain three provider reads per input position at
256, 1,024, and 4,096 elements, including the lazy-branch case. K=3 workloads
also measure the wider product accumulator.

The [independent verifier](verify.py) recomputes raw medians, checks report
values, verifies source/toolchain/lockfile identity, and checks balanced
allocation and retained evidence hashes. Its [receipt](root-verification.json)
also verifies all seven gates and reproduces both exact-rational datasets.
The [performance plan](performance/PLAN.md) describes the selected workloads
and the limits of these measurements.

The [independent review](review.md) accepts the bounded implementation and
records a performance limitation: a large literal expression nested below a
non-cacheable arithmetic branch can repeat cache-classification work for each
projected cell. Reference-read scaling for the measured fixtures does not
establish an asymptotic bound for every expression tree.
