# ODS paired-statistics resource and cache review

This review covers the resolver-aware and scalar implementations of `CORREL`,
`COVAR`, `PEARSON`, `RSQ`, `SLOPE`, `INTERCEPT`, `STEYX`, and `FORECAST`.
It applies ADR 0005's bounded work and storage rules and ADR 0006's typed
failure, source-fence, and publication rules to the local paired-statistics
contract. The review changed no production or test source.

The disposition is **PASS** for the frozen resource and demand-cache
boundaries. The implementation streams paired references, retains only fixed
numeric state, preserves typed failures, and reuses only invariant forecast
fits.

The kernel exposes fixed charged work units for an admitted pair,
finalization, fit preparation, and a forecast query. Those are coarse logical
budget units over finite-width arithmetic; they are not an instruction-level
count of every limb or bit operation inside restoring division. Cancellation
is checked at the surrounding charge boundaries. This review makes no CPU
throughput, latency, or RSS claim.

## Frozen source identity

The frozen baseline is commit
`aa48eee68cb6ee0904523394d4ff8015dce6e595`. Complete custody, including the
profile source set, is recorded in [`gates/freeze.json`](gates/freeze.json).
The selected paired source and test inputs reviewed here have these SHA-256
values:

| Frozen file | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `5dd68e5e86192dfd7b0f1dae7d72a6dc2f1fcfe89178b3dc9f50c8605441fca4` |
| `crates/litchi-ods/src/codec/formula/evaluation/descriptive/moments.rs` | `72c430fa07be48a5119bf2fbd70a2552d4500ae3c9e18cf72933b508da4d1b5d` |
| `crates/litchi-ods/src/codec/formula/evaluation/dyadic.rs` | `9509ed7613b305ee6ae8f7e7aa6136ad88a5450081f0f231ff0079fafce9a411` |
| `crates/litchi-ods/src/codec/formula/evaluation/paired.rs` | `37463db442f2e94cf7b590f8ce4c10e76f9a1e23a094cda2baa53c8a0878fe8f` |
| `crates/litchi-ods/src/codec/formula/evaluation/paired/kernel.rs` | `ef17f203fa0aba0c93a6bc655aa722319a2b4590c94becffc66a406003d24257` |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `54631be0e2dfc5bbbff6a32ecedff2c87f48bfa19412f2051e9c0dbb4df6ecf0` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/paired.rs` | `3e884a1aa675e114f1f65b9e59d05b7b029a3a7fe67371d24893692dfa854fa7` |
| `crates/litchi-ods/tests/ods_formula_paired_evaluation.rs` | `1c60294138922b59a6d3c68822e43b26d04c3f5ac5bf5f37bdde18f621e625a8` |
| `crates/litchi-ods/tests/ods_formula_paired_limits.rs` | `2a15e5580346c0f633cb045103409f793c37989ee46e0ea518b78b587cfddb2e` |
| `crates/litchi-ods/tests/ods_formula_paired_native.rs` | `ab3953ba42151c623319a563c03003c8e844cb0c01dba022c91e284d39486c17` |
| `crates/litchi-ods/tests/ods_formula_paired_oracle.rs` | `8c5251970f36a94819baf5f714f756c0eb03e8f3d7ee5bd9f3feff09e1ba175a` |
| `docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/contract.md` | `9ce79b7b5014c6e29c1d30c61a77187d6dc4bce9be7810e1f5006e8ef147aadd` |
| `docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/numeric_oracle.py` | `03ea2bdff5ba6aca8d5d8ed8775dd479e517bd2a73c2d0c617b968d6360c164e` |
| `docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/numeric-goldens.json` | `038d8b16be3fcf8be2a6234a35ac91d5b5a819ba1b512a4dbcbcf7b802f92f2a` |
| `docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/native/cached-results.json` | `6907eb0fde9145f256587a2aa4d07018221ffc407fd6192acee452446446876e` |
| `docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/native/provenance.json` | `82ee3b3fba3f6f2762b56e64c4ed37c4982a9e5060cfb2516787684bb7e5aa49` |

The isolated gate checkout uses the frozen `Cargo.lock` with SHA-256
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`, as
recorded in the freeze manifest. The ambient root lockfile is a different
workspace copy and is not the custody input for this review.

## Resource and ownership checks

* **Shape and read admission:** `data_shapes` and `shape_error` validate the
  two complete ForceArray descriptors, checked cell counts, and orientation
  before `scan_pairs`. Known 3-D, multi-area, and `ReferenceList` data shapes
  are refused before resolver reads. `FORECAST` rejects a known query list
  before the data scan. A computed expression may first have to produce its
  descriptor; the reducer never scans a descriptor that has already been
  rejected. The RSQ unequal-count subtype is selected before reads as
  required by the contract.
* **Streaming references:** `scan_pairs` visits the aligned pair in source
  order. `read_member` charges cell work before a reference read;
  `read_reference_cell` checks the cumulative reference limit and cancellation
  before the provider call, checks cancellation again after it, and increments
  successful reads only after the provider returns. The cell count is shared
  across both data arrays and all accepted occurrences. No range-sized cell,
  text, or pair vector is constructed.
* **Admission and text:** Paired data admits finite `Number` values and omits
  `Empty`, `Text`, and `Logical` members positionally. Formula errors are
  retained; missing, unsupported, complex, and non-finite members become
  generated formula errors. `read_to_element` preserves provider Text as a
  borrowed value. Reference and inline Text bytes are charged even when the
  aligned member is omitted, and no per-observation text clone is made.
* **Bounded numeric state:** `ExactPaired` retains fixed-width dyadic sums and
  centered workspace (34-, 67-, and 160-limb compile-time containers), while
  `PairState` retains only fixed error records. `ForecastFit` is a fixed-size
  exact fit payload. Arithmetic overflow is mapped to the reducer's generated
  numeric error; resolver, allocation, resource, cancellation, and source
  failures remain typed evaluator failures.
* **Work and cancellation:** The value path charges the fixed pair operation
  before `push_pair`, and charges finalization or fit/query work before the
  corresponding kernel call. The constants bound the selected finite-width
  logical operations, including the fixed arithmetic widths and quotient
  envelope. `cmp_shifted` now compares an allocation-free shifted limb view,
  so restoring division does not materialize a shifted temporary or rescan
  every bit for each quotient position. The charges still should not be read
  as a literal primitive-operation count; the frozen forecast query charge is
  a coarse logical unit for a bounded restoring-division call. There is no
  uncharged unbounded loop or allocation, and the review makes no performance
  claim.
* **Allocation and drop order:** Query matrix output is allocated through the
  checked element-vector reservation. The evaluator-local forecast cache uses
  checked capacity growth and storage-budget reservation; replacement and
  insertion paths retain the old reservation until the new capacity is safe,
  and the reservation is declared after the vector so values are dropped
  before budget release. AST arrays, reference metadata, and output geometry
  retain their existing checked limits.
* **Error precedence:** `PairState` records the first formula and generated
  errors by source argument and ordinal, but `scan_pairs` continues all
  admitted cells after an error. Later typed provider, unsupported, resource,
  cancellation, allocation, or source-version failures return as
  `EvaluationFailure`; they are not caught by the reducer or converted to a
  formula error. Formula-error publication waits for the required scan and
  final source/cancellation checks.
* **Source fence and publication:** The outer value evaluator checks source
  version and cancellation before evaluation and again after fit/query work,
  immediately before publication. A source change or cancellation therefore
  supersedes a retained formula result.

## Demand-cache checks

Paired functions are classified as sequence operations before projected-branch
cache lookup and use the paired apply path. Matrix-argument scheduling marks
both data arguments as complete ForceArray descriptors. For `FORECAST`, the
data arguments remain complete while the ordinary query stays
position-sensitive; an explicit Array query keeps its complete shape. The
shape and reference-kind paths preserve this data-versus-query distinction.

The forecast fit cache is created only when the complete data arguments pass
the conditional-criterion cacheability walk. It stores the exact fit and its
separate formula/generated error payload, then evaluates each query coordinate
at its own position. A cached fit is never used to suppress a query conversion,
resolver read, cancellation, or typed failure. Typed failures are returned
before cache publication and are never represented as formula-error payloads.
The cache is evaluator-local and source-fenced; only the fixed fit or scalar
error payload is retained.

Nested demand propagation marks paired data arguments complete while keeping a
computed scalar query position-sensitive. `MUNIT` remains excluded from the
generic full-argument criterion walk: a direct MUNIT matrix input uses the
function's first-cell input invariant, while a MUNIT expression consumed as a
conditional scalar criterion remains position-sensitive. These are distinct
input/output cases and are not conflated by the paired cache path.

The seven non-`FORECAST` functions publish one scalar result per invocation.
`FORECAST` obtains its matrix output shape from its query argument and reuses
one invariant data fit; it does not reread or refit the data for each output
cell.

## Validation evidence

The frozen focused targets pass with zero failures or ignored tests:

* `ods_formula_paired_evaluation`: 12 tests;
* `ods_formula_paired_limits`: 9 tests;
* `ods_formula_paired_oracle`: 8 tests; and
* `ods_formula_paired_native`: 2 tests.

The focused limits cases cover zero-read shape/list refusal, cumulative
reference limits, work and cancellation boundaries, typed provider failure,
formula-error continuation, borrowed Text accounting, source fencing, and
invariant FORECAST fit reuse. The retained independent oracle contains 464
observations over 39 fixtures. The pinned native evidence retains 42 finite
observations across all eight functions and remains corroborating host
behavior rather than the semantic authority.

The isolated gate receipts record 1,518 passing package tests with zero
failures or ignored tests, plus passing all-target Clippy
with warnings denied, rustdoc with warnings denied, package and batch format,
crate-boundary validation, and `git diff --check`. The package log has zero
failed and zero ignored tests. The gate artifacts and source-stability hashes
are held under `gates/`; no timing, allocation-count, RSS, or throughput
claim is made here.
