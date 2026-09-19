# ODS descriptive-statistics resource and cache review

This review covers `AVEDEV`, `DEVSQ`, `GEOMEAN`, `HARMEAN`, `KURT`, `SKEW`,
and `SKEWP` in the scalar and resolver-aware formula evaluators. It applies
ADR 0005's bounded work and storage rules and ADR 0006's typed-failure,
source-fence, and publication rules to the local descriptive-statistics
contract. The review changed no production code or tests.

The disposition is **PASS** for the resource and demand-cache boundaries. The
frozen implementation has no remaining resource blocker.

## Frozen source identity

The baseline is commit `b8e5d5fe257fd95747c69a3c44a53cedd96f77ed`. Source
custody is recorded in [`gates/freeze.json`](gates/freeze.json). The selected
files reviewed here have these frozen SHA-256 values:

| Frozen file | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `500d71a8a6f6d7776011d04d4cc462e4299a81d8a8b60916b7a8624244d91cda` |
| `crates/litchi-ods/src/codec/formula/evaluation/descriptive.rs` | `0311f15ba569cbfe52b75a67bdc300b3dc74e0ca98db8d401f2e14541253ed3d` |
| `crates/litchi-ods/src/codec/formula/evaluation/descriptive/harmonic.rs` | `8e273a095fbd8bca95be1a5778a620eb31847d8e98b4c5559cebe26bb1704cab` |
| `crates/litchi-ods/src/codec/formula/evaluation/descriptive/moments.rs` | `e857b413183796793057ae59280d37579eec26e6f6ac15a5f3275f7d83e706d4` |
| `crates/litchi-ods/src/codec/formula/evaluation/descriptive/reciprocal_sum.rs` | `a05e30c707e8865cb72bc6edd81740c03edb9222ae4df94e150f4227e0a31920` |
| `crates/litchi-ods/src/codec/formula/evaluation/numerics.rs` | `627ec3d9c5f0f15d43f483fc64b28a861c9235877173a1db087a82c7697ee6e0` |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `2d230ccb34781b985c430964efd5fceb86ea25bbdc3dc1ebbd14efc516928be6` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/descriptive.rs` | `6ef63b74468e40dba853b81b9b2d0ba692a14c6df39d2ab6209c893aacef9c8e` |
| `crates/litchi-ods/tests/ods_formula_descriptive_evaluation.rs` | `358893f8efe6bdd96032d35768832ae89ccf2d680b2457bd88ac8ee8859dea82` |
| `crates/litchi-ods/tests/ods_formula_descriptive_limits.rs` | `7b06dcc3da84cd4fc5f4409d4a981ef5a1044a120b846b7c9b685ce7ddc60e61` |
| `crates/litchi-ods/tests/ods_formula_descriptive_native.rs` | `f79d29dac020a958369cddc103cf624295dfcbdcb6022ae62de79d9ee34ed148` |
| `crates/litchi-ods/tests/ods_formula_descriptive_oracle.rs` | `215680a9981ab7a51c23a7a8b9b11ed80b3c9176a946735c53a43cd002c24bbe` |
| `contract.md` | `6d0127bbf1867553860ba20013aab530c9931ea961e022532e9c16df4424965f` |
| `numeric-goldens.json` | `d59ac0779387ff5cb56de44f21d6760e57a82b30d8ac9fc6c6a253c235a2dfb9` |
| `native/cached-results.json` | `3eeb69e01c6885fbb86d7a7b4d4d8bb23c43eb88bf501b37e16ce20a617112bf` |
| `native/provenance.json` | `57cf56e54ea1a44f621d788b63790e6c5d92e4fc67bb2d3ba49dd0b21a1eff1d` |

The isolated gate checkout uses `Cargo.lock` SHA-256
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`, as
recorded by the freeze manifest. The ambient workspace lockfile is a
different copy and is not the gate dependency input.

## Resource and ownership checks

* **Geometry and admission:** Area, sheet-span, row, column, cell, array, and
  descriptor products use checked arithmetic before a scan. Cumulative
  reference reads and retained metadata are bounded across all arguments and
  all list occurrences. AST arrays, runtime descriptors, and cache scratch
  use fallible capacity under the evaluator storage budget.
* **Streaming reads:** The value reducer walks each retained area in source
  order, then sheet, row, and column order. Cell work is charged before every
  resolver call. `read_reference_cell` checks the cumulative read limit and
  cancellation before the provider call, checks cancellation afterward, and
  increments successful reads only after the provider returns. No reference
  cell is copied into an input vector.
* **Function-specific state:** `DEVSQ`, `KURT`, `SKEW`, and `SKEWP` use one
  admitted scan through fixed-width exact dyadic moment state. `AVEDEV` keeps
  its complete descriptors and performs a charged replay for absolute
  deviations. `GEOMEAN` uses fixed sign, exponent, and ratio state.
  `HARMEAN` uses a bounded reciprocal first pass and enters an exact replay
  only when its interval cannot classify the denominator or final quotient.
  The exact fallback retains aggregate rational limbs only; it never retains
  a cell or text vector.
* **Bounded arithmetic footprint:** The moment helper uses 160 limbs for raw
  sums and 384 limbs for calculation scratch. On the frozen target the fixed
  `ExactMoments` state is about 5.2 KiB, the AVEDEV replay state about 6.2 KiB,
  and the combined descriptive kernel about 11.5 KiB. These are fixed-size
  states, independent of reference length. Harmonic vectors are the only
  adaptive aggregate storage: their count-derived limb ceiling is checked,
  actual capacity growth is preflighted, and old-plus-new allocation overlap
  is included before a `Vec` replacement.
* **Storage and work boundaries:** Harmonic reservations grow by the actual
  simulated capacity delta rather than reserving the count-derived maximum.
  The reservation covers all four aggregate vectors, with the fourth vector
  serving as final numerator scratch, and is declared so aggregate vectors
  drop before the reservation.
  Exact harmonic pushes and finalization precharge work from occupied limb
  precision plus the binary64 exponent-span bound before mutation. A failed
  `try_reserve_exact` remains a typed allocation failure; a budget refusal
  remains a typed memory resource failure. All temporary reservations are
  released on success, formula-error publication, and typed failure.
* **Formula errors and typed precedence:** The first formula Error in
  conceptual argument/reference order is retained while the primary scan
  continues. A later resolver `Unsupported`, cancellation, source change,
  work or storage limit, or allocation failure remains an
  `EvaluationFailure`; it supersedes the retained formula result and is not
  converted into a formula Error, cached, or caught by `IFERROR`. Scheduled
  AVEDEV and HARMEAN replay passes retain the same precedence and cumulative
  budgets.
* **Text and fences:** `read_to_element` enforces the text limit and keeps
  provider Text borrowed. Text-byte work is charged without cloning one
  string per cell. The value evaluator checks source identity and
  cancellation before evaluation, after all required passes, and immediately
  before publication.
* **Shape refusal:** `DEVSQ` and `SKEWP` reject an already-retained
  `ReferenceList` descriptor before `scan_reference`, returning `#VALUE!` with
  zero resolver cell reads. Direct scalar and literal-array paths likewise do
  no resolver reads. A computed expression can read while producing its
  descriptor; the zero-read guarantee applies after the refused shape is
  known.

## Demand-cache checks

The descriptive family is classified as a sequence operation before projected
branch cache lookup and uses complete matrix argument descriptors. A retained
Reference or ReferenceList therefore stays complete under a projected lazy
`IF`, and a two-pass AVEDEV operation reuses its descriptors rather than
implicitly intersecting or rereading one output position. Computed reducers
remain conservative unless their complete sequence is proven invariant.

The cacheability walk propagates complete context through descriptive sequence
arguments and nested statistical criteria. Scalar parameters remain
position-sensitive. `MUNIT` stays excluded from full-argument propagation, so
its scalar criterion is evaluated at the requested coordinate while a direct
matrix input keeps its first-cell invariant. Shape and reference-kind planning
classifies these seven functions as scalar reductions independently of their
sequence geometry. Cache entries contain only scalar Number or formula-error
payloads with source/context identity; typed
evaluator failures and borrowed text are never cached.

## Validation evidence

The frozen focused suites passed 12 semantic evaluator tests, 18 resource and
typed-failure tests, 2 native-profile tests, and 7 independent-oracle tests.
The resource target includes the 20,001-row repeated reciprocal cancellation
case under a 64 KiB storage cap, cumulative replay read limits, source and
cancellation fences, typed provider failures, zero-read shape refusal, and
reservation refund checks. The oracle covers 413 observations across 54
fixtures; native evidence retains 39 observations across the seven functions.

The isolated package gate passed 1,477 tests with zero failures and zero
ignored tests. The seven frozen gates passed: package tests, all-target
Clippy with `-D warnings`, rustdoc with warnings denied, package formatting,
selected-file rustfmt, crate-boundary validation, and `git diff --check`.
The gate receipt and source-before/source-after equality are recorded under
[`gates/`](gates/). This review makes no timing, throughput, RSS, or allocator
performance claim.
