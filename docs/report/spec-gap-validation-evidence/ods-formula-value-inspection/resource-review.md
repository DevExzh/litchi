# ODS value-inspection and conversion resource review

This review covers the sixteen §6.13 functions `ERROR.TYPE`, `ISBLANK`,
`ISERR`, `ISERROR`, `ISEVEN`, `ISLOGICAL`, `ISNA`, `ISNONTEXT`, `ISNUMBER`,
`ISODD`, `ISTEXT`, `N`, `NA`, `NUMBERVALUE`, `TYPE`, and `VALUE`. It applies
ADR 0005's bounded work and storage rules and ADR 0006's typed-failure,
source-fence, and publication rules to the value-inspection contract. This
review changed no production code or tests.

The independent disposition is **PASS** for the resource and demand-cache
boundaries. The frozen source closure matches the freeze manifest. Independent
source verification reports stable source-before/source-after hashes, all
seven required integration gates passing, and 1,623 package tests with no
failures or ignored tests. Optional performance capture remains pending.

## Frozen source identity

The frozen baseline is commit
`d623f3c2ecc0c837017f700174656f0e443759a5`. Complete custody is recorded in
[`gates/freeze.json`](gates/freeze.json), which contains 82 selected files.
The relevant inputs reviewed here have these SHA-256 values:

| Frozen file | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `b88159d778793e14583b5807bfc6ccdff9877eb59a19b6cdbc3514120db80392` |
| `crates/litchi-ods/src/codec/formula/evaluation/inspection.rs` | `594fa0a8110296259532e20e7901685eeea9502eb8927f2915d58bfdf52d8918` |
| `crates/litchi-ods/src/codec/formula/evaluation/inspection/parse_value.rs` | `f29a42e77912a7dd65acc92a82f306fd9875e7d6a3410974029e940276ebf504` |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `9d30e3f77ad8b75a403d126c87259f7637041db52ede7441157c649bdec60e8d` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/inspection.rs` | `4dfc0cf96e6bcd963906143b0ad07445e6883362073bb557a0fb6a139fff1a14` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/scalar.rs` | `cd882fa914bffa26ede20c51a4483c53fc36f4070e37cfbfef2bfec6e2db69f3` |
| `crates/litchi-ods/tests/ods_formula_inspection_evaluation.rs` | `7e882c07b6a6fb1ab1495d3d6ecbb5d23258728e33e6820ccdecb290e03ab8e1` |
| `crates/litchi-ods/tests/ods_formula_inspection_limits.rs` | `5a1091d2b32613658351a380a44bb5cc9a6da6b0abc43c758e4e36e280e96a1e` |
| `crates/litchi-ods/tests/ods_formula_inspection_oracle.rs` | `eefb80c1e06a9101b5e161fd28ea95f76c28f0a1c5bb4c3c5efd65d425f8b0d5` |
| `contract.md` | `f227e861c55e3360d60d41a42ee7f101ba3c29bbc00d25d56cc01d80c3cd7923` |
| `inspection-goldens.json` | `4cd3cfd4da0df1cfef9aa0d463b0c9b369c2c94579597a8a0ddc93df81ba1aa4` |
| `native/native-results.json` | `523956c654c2c6cede4cc7900f81e7526a166de0db7a53afeb02b5e1dee1afba` |
| `native/provenance.json` | `420aba22f97db588468e4d768d014a2d15b2ca4336b081b00b27c2683b6fd996` |

The isolated gate checkout uses the retained gate `Cargo.lock` with SHA-256
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`, as
recorded in the freeze manifest. The ambient workspace lock has a different
hash and is not the dependency input for the gate custody claim.

## Resource and ownership checks

* **Shape and type admission:** The value inspection bridge validates arity
  and rejects a known `ReferenceList` before selecting a cell. `TYPE` handles
  an admitted rectangular Reference as a complete scan and rejects a list
  without scanning it. Known invalid `NUMBERVALUE` separator pairs are
  detected from scalar descriptor values before source-cell reads; matrix
  output is bounded and filled with formula `#VALUE!` elements. The stated
  zero-read boundary begins once a descriptor is known, so a computed
  reference expression may have performed its own upstream work.
* **Streaming reference work:** `TYPE` visits every retained area and every
  sheet/row/column cell, discarding each completed value immediately. The
  elementwise predicate and conversion paths select one cell per output
  coordinate and keep the reference descriptor instead of materializing a
  range. Cell work is charged before selection/provider access.
  `read_reference_cell` enforces the cumulative read limit and checks
  cancellation before and after the resolver call; successful-read accounting
  happens only after the provider succeeds.
* **Formula-error and typed-failure precedence:** Raw predicates,
  `ERROR.TYPE`, and `TYPE` inspect formula errors as values. `N`, parity,
  `NUMBERVALUE`, and `VALUE` propagate formula errors. A complete `TYPE`
  reference scan continues after a formula error, and a later unsupported,
  resource, cancellation, source, or provider failure remains a typed
  `EvaluationFailure` that supersedes the formula value. `IFERROR` and the
  inspection functions do not catch or rewrite those typed failures.
* **Borrowed text and parser bounds:** Resolver text is checked against the
  text-byte limit and retained as a borrowed `TextValue`. `NUMBERVALUE`
  charges source and separator bytes, uses checked source-length arithmetic,
  reserves bounded normalization scratch, and checks execution during long
  normalization, lexical validation, and percent handling. `VALUE` charges
  input bytes and uses a checked, budgeted scratch buffer for long grouped
  input; short parsing remains fixed-size. No reference cell or complete
  range is cloned merely to inspect or convert it.
* **Output and reservation bounds:** Shape products, reference descriptors,
  AST traversal frames, demand-cache storage, parser scratch, and matrix
  output arrays use checked limits and fallible reservations. Output vectors
  are paired with their reservation so typed, cancellation, and formula-error
  exits drop the vector before releasing its storage lease. `TYPE` retains no
  cell vector, and `NA` has constant-size state.
  Complete `Any`/`Array` argument contexts add one budgeted saved-context entry
  and two bounded VM frames per argument. The saved mode, current position,
  projection, and matrix continuation are restored through the flat VM; on an
  early return, probe-owned frames, values, contexts, and their reservations
  are dropped before the suspended context is resumed.
* **Fences and publication:** The outer value evaluator checks cancellation and
  the source version before evaluation, checks the source again after the VM
  run, and checks cancellation immediately before publishing the result.
  Resolver text remains valid through the shared expression/result lifetime.

## Demand-cache checks

Inspection functions are scheduled through the value VM's matrix argument
path. `TYPE` is classified as a sequence-like scalar result before projected
cache lookup; a complete direct reference or array descriptor may be scanned
once and cached only as a scalar number or formula-error payload. The cache
path does not retain arrays, references, borrowed text, or typed failures.

`VisitCompleteArgument` clears a projected cell demand while preserving the
caller output position. This keeps a computed `Any` array complete for `TYPE`
and keeps a direct `N(reference)` intersection tied to the requested position.
`N` is deliberately absent from the sequence pre-cache and full-argument
propagation sets, so a position-sensitive scalar result cannot be reused at a
different output coordinate. The scalar inspection bridge restores an omitted
AST argument as `Input::Missing` after popping its stack-balancing placeholder;
the optional `NUMBERVALUE` slots therefore use defaults without another AST
walk, source read, or heap-backed argument vector. Direct `MUNIT` keeps its
first-element profile, and its scalar size argument remains excluded from
complete-reference propagation.

`ISBLANK`, the other predicates, `N`, `VALUE`, and `NUMBERVALUE` remain
position-sensitive in projected evaluation. Their matrix selectors preserve
the relative output coordinate, including a 3-D current-plane selection and
the typed-later-cell `TYPE(MUNIT(N(reference)))` case. The classifier rejects
computed inspection/text descendants in invariant matrix branches; this is a
conservative reread tradeoff that preserves correctness. Direct `MUNIT`
matrix arguments retain the existing first-cell profile, while its scalar size
criterion is excluded from complete-reference propagation.

The shape planner and reference-kind planner classify `TYPE` and `N` as scalar
result operations while retaining their distinct complete-scan and
intersection semantics. Reference-list refusal is performed before the
elementwise mapper or generic materializer can read cells.

## Validation evidence

Against the frozen tree, the focused targets pass:

* six inspection evaluation tests;
* five inspection resource and typed-failure tests; and
* one independent-oracle integration test.

`cargo check --locked --offline -p litchi-ods` passes. The related regression
targets also pass in the current tree: 24 matrix evaluation tests, 12
statistical evaluation tests, and 22 conditional evaluation tests. The
retained frozen gate receipts under [`gates/`](gates/) show all seven required
checks passing: package tests, all-target Clippy with `-D warnings`, rustdoc,
package formatting, and the selected-file rustfmt check. The isolated receipt
covers 1,623 package tests with no failures or ignored tests. `freeze.json`
selected-file hashes match the current source closure; its isolated Cargo.lock
is `58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`.
The ambient workspace Cargo.lock differs and is excluded from the frozen gate
claim. Optional performance capture remains pending and is outside the
required correctness and resource gate disposition.

No throughput, allocation-rate, RSS, or timing claim is made by this review.
