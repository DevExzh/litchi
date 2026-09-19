# Independent ODF 1.4 numeric aggregate review

## Disposition

The frozen implementation satisfies the reviewed semantic contract for
`SUM`, `PRODUCT`, `SUMSQ`, `SUMPRODUCT`, `SUMX2MY2`, `SUMX2PY2`, and
`SUMXMY2`. I found no remaining semantic, numeric, resolver, or resource
accounting blocker. The review accepts the batch for the bounded scalar and
resolver-backed value profiles described in
[`contract.md`](contract.md).

The final isolated gate receipt reports 1,238 ODS tests passed, with zero
failures and zero ignored tests. All seven required gate commands passed and
the source-before and source-after manifests are identical. The review did
not change production code or tests.

The performance harness was reviewed for custody and workload coverage, but at
review time no timing review had been performed. This report therefore makes
no throughput, allocation, or resident-memory claim. The harness's measured
guarantees belong in a separately retained profile receipt.

## Source and identity

The normative source and profile decisions are recorded in
[`contract.md`](contract.md). It cites the local OpenFormula 1.4 Part 4
archive (SHA-256
`9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`) and
the extracted Part 4 member (SHA-256
`ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`). The
final contract hash is
`fa05c4bc7f81d927330695a27f9da35954aa37ab4d015cf3b3a026b1c1fa37d6`.

The frozen source identity is in
[`gates/freeze.json`](gates/freeze.json); the gate manifests are
[`source-before.json`](gates/source-before.json) and
[`source-after.json`](gates/source-after.json). Their SHA-256 is
`394a6d4d508c66a685e43c1620b268b2a8dfd60a566d751aa05bb22663456f25`.
The final gate outcome is recorded by
[`results.json`](gates/results.json) and
[`verification.json`](gates/verification.json); the latter reports
`stable_sources: true` and `all_required_checks_passed: true`.

## Semantic review

The dispatch and evaluator boundaries match the contract. The scalar bridge
implements scalar NumberSequence conversion and the 1×1 instance of each
ForceArray function. The value bridge keeps full arrays and references in
ForceArray mode even when the surrounding evaluation is scalar. `SUM` and
`PRODUCT` flatten ordered ReferenceLists, while `SUMSQ` accepts one Reference,
including the tested single 3-D Reference, and rejects an explicit
ReferenceList. Referenced Text, Empty, and distinguishable Logical cells are
omitted from NumberSequence functions; inline arrays use the documented
element conversion. Matrix Text, Logical, Empty, Error, Missing, and
unsupported-cell behavior is distinct and tested.

`PRODUCT()` remains a strict zero-argument arity error. A supplied sequence
that selects no Number cells uses product identity one. `SUM()` and `SUMSQ()`
use additive identity zero. Matrix functions require equal rectangular
shapes, do not broadcast, reject multi-area ForceArray inputs, and preserve
the scalar 1×1 case. Formula errors are tracked in source argument order and
row-major cell order even when the implementation reads corresponding
matrix coordinates together. Typed resolver, cancellation, source-version,
work, storage, and resource failures remain evaluator failures and cannot be
hidden by formula error handlers.

The numeric implementation is aligned with the final profile. `SUM` uses a
fixed binary accumulator. `PRODUCT` keeps sign, normalized mantissa, and
exponent separately and rounds its normalized binary64 mantissa at each factor
update, delaying final range conversion. `SUMSQ` and all three `SUMX*`
functions use the fixed wider pair/square accumulator. `SUMPRODUCT` uses the
additive path for K=1, exact fixed-width pair products for K=2, and a fixed
128-limb K>2 signed window: 8,192 bits total, 64 carry bits, and an 8,128-bit
admitted input span. The K>2 limb reduction is exact for the per-factor-rounded
53-bit product terms; it does not silently drop tails. A span refusal is a
typed `Resource::Memory` result with required and profile-limit bytes in both
the scalar and value bridges. Caller budgets remain additional limits.

Final conversion uses the documented round-to-nearest-even behavior, including
subnormal preservation and the distinction between a finite rounded maximum
and a non-finite final conversion. Intermediate product or square overflow is
not treated as final overflow, allowing cancellation vectors to complete.

## Validation evidence

The focused suites and retained independent data provide the following
coverage:

| Evidence | Result |
| --- | --- |
| Scalar aggregate evaluation | 8 passed |
| Value array/reference aggregate evaluation | 13 passed |
| Native fixture receipt | 2 passed; 48 selected cache observations |
| Exact numeric oracle | 2 tests; 17 targeted + 336 seeded observations |
| ODS library unit tests | 399 passed |
| Final locked ODS package tests | 1,238 passed, 0 failed, 0 ignored |
| Clippy, rustdoc (`-D warnings`), format, batch format, boundaries, diff check | all passed |

The exact rational observations are retained in
[`numeric-goldens.json`](numeric-goldens.json) and
[`differential-goldens.json`](differential-goldens.json). The oracle test
evaluates the public value API over closed literal expressions, not just a
helper-level arithmetic function. The targeted vectors cover wide
cancellation, pair-square cancellation, subnormal results, intermediate
overflow/underflow, and K>2 cancellation. The fixed-window refusal is covered
by the authored value-array harness test; it is not counted as a Fraction
oracle observation.

The native receipt in [`native/`](native/) contains 48 selected observations
covering all seven functions: SUM 7, PRODUCT 9, SUMSQ 5, SUMPRODUCT 9, and
six for each SUMX* function. It excludes the zero-argument native PRODUCT row
because the selected profile is strict, a SUMPRODUCT formula-cell input, and
the SUMPRODUCT row whose literal `"Unknown"` is treated as zero by the native
cache while this contract reports malformed forced-array text as `#VALUE!`.
These are explicit compatibility variances; no cached formula result is used
as an input.

## Harness and regression review

The value evaluator's iterative matrix cache classifier now admits fixed local
references for aggregate, MUNIT/TRANSPOSE, and database branches while still
rejecting source references, names, labels, and automatic intersections. The
full package gate and the existing matrix/database/complex test suites passed,
so no regression was observed in those paths. The aggregate tests exercise
scalar projection, lazy branches, nested aggregates, paired references,
ReferenceLists, 3-D sequence traversal, cancellation, and bounded work and
storage failure paths.

The performance harness source and plan were inspected. It uses fresh child
processes, three warmups, fifteen samples, matched controls, candidate-only
aggregate cases, resolver read counts, budget accounting, checksums, and
source/toolchain/lockfile custody. Its profile input list includes the final
contract, numeric goldens, native receipt, and their generators. At review
time these properties are a harness review only; no timing conclusion is
drawn here.

One measured-performance limitation remains explicit: `apply_function` calls
`cacheable_scalar_branch` before looking up a demand-cache entry. A large
literal AST nested below a non-cacheable arithmetic branch can therefore
repeat classifier work for each projected output cell. The current tests prove
bounded refusal for deep ASTs and bounded reference reads for the supplied
scaling cases; they do not establish an all-AST asymptotic bound. This does
not change aggregate semantics or create a correctness blocker, but it must
remain visible in any later performance report.
