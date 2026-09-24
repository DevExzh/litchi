# ODS order-statistics resource and cache review

This review covers the resolver-aware and scalar implementations of `MEDIAN`,
`MODE`, `LARGE`, `SMALL`, `PERCENTILE`, `PERCENTRANK`, `QUARTILE`, and `RANK`.
It applies ADR 0005's bounded work and storage rules and ADR 0006's typed
failure, source-fence, and publication rules to the local order-statistics
contract. The review changed no production code.

The disposition is **PASS** for the resource and demand-cache boundaries. The
frozen implementation has no unbounded range materialization, typed-failure
conversion, or position-insensitive scalar cache path.

## Frozen source identity

The baseline is commit `d6c0485365644a5525b7302ba0b3666d43df1b1f`. The selected
source custody is recorded in [`gates/freeze.json`](gates/freeze.json). The
following hashes are the frozen inputs reviewed here:

| Frozen file | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `4a7a3ec13a425ed44432a220a8e14c75c8781934cdce6b9632b085a5e9967c1d` |
| `crates/litchi-ods/src/codec/formula/evaluation/order.rs` | `516cc2f144de18522ef40f708e542573a259287aff9cfa0741c04beb3f86d86a` |
| `crates/litchi-ods/src/codec/formula/evaluation/rounding.rs` | `82a46b6ba227a79979f284e9fbaf9838626fe2a1cad3bafdac968c77a67a6d03` |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `190609b5036a3a64a0ac035f9956e215dbb4c20d3152257441e6ed803f7a46e3` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/order.rs` | `9f66676154ed7945f02d7e4cdb321d9c5ae9d5e8c61ef3d2cf9ab991a288a88b` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/matrix.rs` | `95ac99a7ded56e77966029f9cc1433aa0b70fe07f1a7f30020589dc973e3a5db` |
| `crates/litchi-ods/tests/ods_formula_order_evaluation.rs` | `8e464969151e6ebf27f3bfd9c7a27932885581c4f7016e3b801914e1d9f183c2` |
| `crates/litchi-ods/tests/ods_formula_order_limits.rs` | `bc9649454de56f84280cbe68c39bdf5ecf1c431659076ba1b00fdbbf6236da5c` |
| `crates/litchi-ods/tests/ods_formula_order_native.rs` | `556a18939c9149c27ea52521e994c510c031087b4858e558512c9e12882eac2e` |
| `crates/litchi-ods/tests/ods_formula_order_oracle.rs` | `b94a6f337ab21a1beec3d4f4a78dacefab8f3651333bb3affb42f9283ea7f6f9` |
| `numeric-goldens.json` | `b3b7eb6b53535c1f55c1f9995f47fbc1c5c8d60c0f0cb55a710a47522c3bbce4` |
| `native/cached-results.json` | `65605b11c2c4e87278ceb2d173177e18c37cc23b6d80b74e657cd148010e9dd3` |
| `native/provenance.json` | `34dff3c78f41143c127523d5d48009efd8db47c426b141074ad757b76d2ee9b1` |

The isolated gate checkout uses `Cargo.lock` SHA-256
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`, as
recorded in the freeze manifest. The ambient root lockfile is a different
workspace copy and was not used for this custody claim.

## Resource and ownership checks

* **Shape and capacity admission:** The value reducer rejects a `ReferenceList`
  for `MODE` and `QUARTILE`, and rejects a list in a scalar parameter, before
  the admitted-capacity pass or a resolver read. The corresponding focused
  tests assert zero reads. `admitted_value_limit` walks all data descriptors,
  checks every rectangle count and cumulative addition, and charges metadata
  work before scanning. The bound is cumulative across variadic arrays and
  reference areas, so two valid arrays do not accidentally share one
  per-array ceiling. Existing array, reference, area, and output descriptors
  retain checked limits and fallible storage reservations.
* **Retained state:** Selection reducers retain one vector of admitted finite
  `f64` values. They do not retain `RuntimeElement`, text, cell, or resolver
  records. `RANK` and scalar `PERCENTRANK` use fixed-size query state when
  their scalar parameters are directly available, retaining only counts,
  boundaries, and bracket information. `OrderValues` grows through checked
  `ensure_capacity` and `try_reserve_exact`; the reservation token is declared
  after the vector, so values are dropped before their budget is released.
  Failed replacement reservations restore the prior token.
* **Streaming reads:** `scan_reference` visits retained areas in their source
  order and walks each rectangle by sheet, row, and column. Cell work is
  charged before every provider call. `read_reference_cell` checks the
  cumulative read limit and cancellation before the call, checks cancellation
  again after it, and increments successful reads only after the provider
  returns. A formula error is retained while the scan continues, allowing a
  later typed resolver, resource, cancellation, source, or allocation failure
  to supersede it.
* **Text and conversion:** `read_to_element` applies the text limit and keeps
  resolver text as `TextValue::borrowed`. Reference Text and distinguished
  Logical cells are filtered by the sequence profile without cloning. Inline
  text conversion charges through the scalar bridge and only the admitted
  number reaches the numeric vector.
* **Sorting and work bounds:** MEDIAN, MODE, LARGE, SMALL, PERCENTILE, and
  QUARTILE sort one retained numeric vector and reuse it for all output ranks.
  The shared sort is iterative heapsort: every comparison and swap goes
  through a fallible work callback, so a standard-library comparator cannot
  hide cancellation or a work-limit failure. Streamed RANK and PERCENTRANK
  charge each candidate comparison and use no sort. Selection, interpolation,
  frequency, and bracket passes are bounded by the admitted vector or stream
  size. Matrix output cells charge cell work; invariant data is not reread or
  resorted for each scalar parameter position.
* **Result and publication boundary:** The value path charges the selected
  reducer completion where a separate result operation is required; MODE's
  frequency pass and the buffered RANK/PERCENTRANK kernels charge their
  comparisons or candidate scans, and matrix output publication charges each
  output cell. The scalar completion paths are constant-size after their
  charged sort or stream. The outer value evaluator checks the resolver source
  version before and after the run and checks cancellation immediately before
  publication, so a cancellation or source change cannot be hidden by the
  final constant-size operation.
* **Typed precedence:** Unsupported values, resource limits, cancellation,
  source changes, and allocation failures remain `EvaluationFailure` values.
  They are not converted into formula errors and are not caught by `IFERROR`
  or cached. Formula errors retain conceptual argument and scan order while
  the admitted scan performs its required reads and work.

## Demand-cache checks

Order functions are classified as sequence operations before projected-branch
cache lookup and use the apply-time cache path. Complete sequence descriptors
enter matrix argument scheduling, so a range used as `Data` remains complete
under a projected lazy `IF`. `LARGE` and `SMALL` retain explicit Array rank
parameters as arrays; scalar rank and query parameters stay in their current
projection and remain position-sensitive. The shape and reference-kind
planners derive output shape from the scalar query parameters, including an
explicit Array parameter for `LARGE` and `SMALL`, while preserving the data
versus parameter distinction.

The criterion walk propagates complete context only for data sequence
arguments and explicit array parameters. Nested order/statistical reducers
therefore reuse invariant sequence descriptors, while a computed scalar
parameter is reevaluated at its requested coordinate. `MUNIT` remains excluded
from full-argument propagation. A direct `MUNIT` matrix input uses its `[0,0]`
cell, including under a projected `IF`; a `MUNIT` expression used as a scalar
conditional criterion input remains position-sensitive at the criterion
coordinate. Cache entries contain only scalar numbers or formula-error
payloads with source and context identity. Arrays, references, typed failures,
and borrowed text are not cached as scalar payloads.

## Validation evidence

The frozen focused suite records 17 order semantic tests, 16 resource-limit
tests, 8 independent numeric-oracle tests, and 2 native-profile tests. The
oracle covers 512 observations over 56 fixtures; native evidence retains 45
observations across all eight functions. The focused resource cases cover
zero-read shape refusal, cumulative reference and storage limits, work and
cancellation boundaries, typed provider failure, formula-error continuation,
projected sequence reuse, scalar parameter position sensitivity, and nested
`MUNIT` criteria.

The final isolated package gate passed 1,421 tests with zero failures and zero
ignored tests. All seven isolated gates passed: package tests, all-target
Clippy with `-D warnings`, rustdoc with warnings denied, package format, batch
rustfmt, crate-boundary validation, and `git diff --check`. The preflight
profile also passed 122 candidate cases and 21 matched controls in both
`evaluate` and `parse-evaluate` phases, with the documented reference-read
bounds. It is an untimed preflight and makes no throughput, allocation, RSS,
or timing claim.
