# Independent paired-statistics semantic and numerical review

Status: **PASS** for the frozen paired-statistics implementation and the
profile in [`contract.md`](contract.md). This review covers `CORREL`, `COVAR`,
`PEARSON`, `RSQ`, `SLOPE`, `INTERCEPT`, `STEYX`, and `FORECAST`. I made no
production or test edits.

## Frozen identity

The source handoff is [`gates/freeze.json`](gates/freeze.json), based on
`aa48eee68cb6ee0904523394d4ff8015dce6e595`. The normative ODF 1.4 archive
SHA-256 is
`9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`; its
Part 4 HTML member is
`ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.
The frozen contract is
`9ce79b7b5014c6e29c1d30c61a77187d6dc4bce9be7810e1f5006e8ef147aadd`.

All selected Rust, test, oracle, native-input, feature-matrix, contract, and
isolated-lock hashes match the freeze manifest. The isolated gate lock is
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`; the
ambient root `Cargo.lock` is a pre-existing different copy and is not the
frozen gate input.

The optimized shared divider/comparator source is
`crates/litchi-ods/src/codec/formula/evaluation/dyadic.rs`, SHA-256
`9509ed7613b305ee6ae8f7e7aa6136ad88a5450081f0f231ff0079fafce9a411`.
Its allocation-free shifted-limb comparator preserves the exact ordering of
the prior bitwise view, including zero, equal-length, cross-limb, and near-
capacity shifts. The added boundary tests exercise those cases.

## Semantic findings

The scalar and resolver-backed evaluators agree with the contract's complete
ForceArray pairing and row-major aligned omission. Finite Numbers form pairs;
Empty, Text, and Logical members omit the aligned position; formula Errors are
retained; unsupported, complex, missing, and non-finite members produce the
selected generated errors. The resolver path streams rectangular references
in lockstep and retains no range-sized input vector.

Shape and pseudotype checks precede reducer reads when descriptors are known.
Equal-shape requirements, scalar-to-matrix refusal, ReferenceList/3-D
refusal, and the RSQ distinction between different checked counts (`#N/A`)
and equal-count orientation mismatch (`#VALUE!`) match the frozen profile.
`STEYX` requires three admitted pairs, rejects zero independent variance, and
publishes exact perfect-fit and constant-dependent results as `+0`.

The Part 4 `INTERCEPT` reference to `LINEST(...;FALSE())` conflicts with the
same section's definition of `Const=FALSE` as a zero constant and with the
description of an ordinary y-intercept. The contract explicitly resolves this
ambiguity as included-constant regression (`ȳ - b x̄`); the implementation
matches that selected profile. This is recorded as a deliberate contract
choice rather than silently treating the contradictory token as a through-
origin regression.

`FORECAST` computes one invariant exact fit for complete data arguments and
converts each scalar query independently. Scalar mode projects the ordinary
query, matrix mode lifts a rectangular query/reference/computed array, and a
produced `MUNIT` array is consumed as an array. The §3.3.2.2.1 exception is
applied to MUNIT's own scalar input; it does not collapse a produced array.
Nested MUNIT scalar criteria remain position-sensitive and excluded from
complete-argument demand-cache propagation.

Formula errors preserve source argument and cell order. An admitted scan
continues after one is retained, allowing typed resolver, resource,
cancellation, allocation, and source-version failures to escape and
supersede a formula result. Query/data formula-versus-generated precedence in
`FORECAST` matches the frozen contract and focused regression.

## Numerical findings

`paired/kernel.rs` retains fixed-width exact binary64 dyadic sums and forms
centered `Sxx`, `Syy`, and `Sxy` before final publication. The fit for
`SLOPE`, `INTERCEPT`, and repeated `FORECAST` queries is precomputed once.
`STEYX` uses the exact nonnegative residual determinant and only reports
`#NUM!` for a genuinely negative radicand. The state is bounded and does not
retain input cells.

The shared divider rounds an exact ratio directly at the binary64 target
quantum, including subnormals, avoiding the former two-rounding case.
Exact algebraic zero is canonical `+0`; a signed nonzero result that
underflows preserves `-0`. Correlation and STEYX use exponent-safe square-root
publication. The independent oracle uses exact `Fraction` operands and a
300-digit `Decimal` square root, with the predeclared finite-publication
tolerance only for sensitive rows.

The retained oracle corpus has 39 fixtures and 464 observations. Reproduce it
from this directory with:

```text
python3 numeric_oracle.py --check
{"observations": 464, "fixtures": 39, "verified": true}
```

Native observations are corroborating evidence only. One retained native
`STEYX` value is exactly 2,005 ULP below the independent mathematical value;
[`native/steyx-deviation-proof.json`](native/steyx-deviation-proof.json)
records the proof and the native comparison remains an explicit host
deviation, not a production tolerance change.

## Validation receipts

The focused frozen Rust targets pass: evaluation `12/12`, resource limits
`9/9`, native comparisons `2/2`, and independent oracle `8/8`. The frozen
isolated package receipt records 1,518 tests passed with zero failures and
zero ignored tests. Strict all-target Clippy, rustdoc, crate formatting, batch
formatting, crate-boundary, and source-diff receipts are all successful in
[`gates/results.json`](gates/results.json), and `gates/verify.py` reports
`verified: true` with all seven gates complete and stable source-before/source-
after hashes.

The independent semantic/numeric disposition is PASS. Root owns the separate
resource/cache report and any performance threshold disposition; neither
changes this semantic review.
