# ODS statistical reducer resource and cache review

This review covers the resolver-aware implementation of `COUNT`, `COUNTA`,
`COUNTBLANK`, `AVERAGE`, `AVERAGEA`, `MIN`, `MAX`, `MINA`, and `MAXA`. It
applies ADR 0005's finite work, storage, source, and cancellation rules and
ADR 0006's typed-failure and publication rules to the contract in
[`contract.md`](contract.md). The review changed no production code.

The disposition is **pass** for the resource and demand-cache boundaries. The
review found no remaining resource blocker in the frozen implementation.

## Frozen source identity

The baseline is commit `f7fe857007b7bcb65e0512c5b5caf9135fe74a5a`. The source
freeze is recorded in [`gates/freeze.json`](gates/freeze.json); the selected
source hashes were re-read after the freeze and matched that manifest:

| Frozen file | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `e8ae9498d114e1abba0ee7825e2f3cbe1afb2c34d838ad2bfec9a2e76bbbe1c5` |
| `crates/litchi-ods/src/codec/formula/evaluation/statistical.rs` | `aa9ac6f994b67cfef1cd93f4171c2a47b67bcba4e24811c49fdf1b8f38238b01` |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `115f552e62884944426a5a93d379d83d31e82a7a4106235c7aa9dfeb4ea097dc` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/statistical.rs` | `ed7efa6d0e406602726484f0b532395f9c265f322a3af7ad7f6e52d63ae8a793` |
| `crates/litchi-ods/tests/ods_formula_statistical_evaluation.rs` | `d5a358d7d0a1793106f8dd22f5ac3f297a6e15d5077190c5a51954f7ac547bc0` |
| `crates/litchi-ods/tests/ods_formula_statistical_limits.rs` | `da90d2fab7890cd43a1194da0bcfb8114e7c76c6510c35923f276479b5706f3b` |
| `crates/litchi-ods/tests/ods_formula_statistical_native.rs` | `00e13c3c1aa3ac484a08c4f527d4662ce127a820f94600d670a6ab2a12effc0f` |
| `crates/litchi-ods/tests/ods_formula_statistical_oracle.rs` | `e709c5280c0c9067fbceecab5a609f3d3294cbdd8344ccb300a68cd823bd6206` |
| `native/cached-results.json` | `cbd13a3e2c05d54d2b782f680773b595272ba89f1a1b7bcede39c13c1f5aec34` |
| `native/provenance.json` | `245e1c2bd6c3c56d6ebb926068aeb55203d58b5ea8cda366b1b02bfaacdf8d0f` |
| `numeric-goldens.json` | `9a2028c4eb3c0bba6442161ddfd227281899763a281f82879720932b98d2f0d9` |

The isolated gate checkout used `Cargo.lock` SHA-256
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`, as
recorded by the same freeze manifest.

## Resource and ownership checks

* **Geometry and storage admission:** `RuntimeAreaSet` checks rectangle cell
  counts, sheet spans, cumulative reference cells, and retained area/record
  counts with checked arithmetic before scanning. AST frames, runtime arrays,
  reference descriptors, and demand-cache entries use `ensure_capacity` with
  fallible reservations and the caller's storage budget. Reducer state is
  fixed-size, including the exact numeric accumulator; reference cells are
  never collected into a vector for a scalar result.
* **Work, reads, and cancellation:** `scan_reference` visits each retained
  area in sheet, row, and column order. Each cell charges work before the
  resolver call. `read_reference_cell` checks the cumulative read limit and
  cancellation before the call, checks cancellation again after it, and
  increments the successful-read count only after the provider returns.
  Array inspection and direct formula-error inspection use the same cell-work
  charging path.
* **Typed failures and formula errors:** The reducer retains formula errors
  according to the contract: `COUNT` ignores them, `COUNTA` counts them, and
  the other reducers retain the first one while continuing the scan. Resolver
  unsupported values, resource limits, cancellation, and source failures
  remain `EvaluationFailure` values; they are not converted into formula
  errors, cached, or caught by `IFERROR`.
* **Source and publication fences:** The value evaluator checks cancellation
  and the resolver source version before evaluation, checks the source again
  after all reads, and checks cancellation once more before publishing the
  result. A changed source or typed failure therefore publishes no retained
  reducer value.
* **Borrowed text:** `read_to_element` applies the per-value text limit and
  creates a borrowed `TextValue`. `COUNTA` and `COUNTBLANK` only inspect the
  borrowed text, while the A variants charge text bytes when mapping text to
  zero; no per-cell resolver string is cloned.
* **Drop order:** Runtime area/record and array structs declare their buffers
  before their reservation tokens, and the iterative cacheability scratch
  vectors declare the reservation before the vector. This drops retained
  elements before releasing their corresponding budget reservation on success
  and typed-failure paths.
* **Zero-read refusal:** `AVERAGE` records `#VALUE!` for an explicit
  `ReferenceList` before `scan_reference`, so the rejected list causes no
  physical cell reads. `COUNTBLANK` enters the resolver scan only for an
  admitted reference. Direct constants and literal arrays require no resolver
  reads; rejecting an already evaluated non-reference descriptor adds no reads.
  Computed arguments can read cells during argument evaluation before this
  type gate, so this is not a zero-read guarantee for arbitrary expressions.
  There is no AST preflight that skips computed argument evaluation.
  Reference-cell and storage-limit tests also verify refusal before provider
  access and release of evaluator memory.

## Demand-cache checks

Statistical functions are included in the projected-branch cache lookup before
their arguments are scheduled and in the apply-time cache get/put path.
Literal array arguments enter `VisitMatrixArgument`; reference arguments retain their area
geometry through `VisitArgument`. Computed arguments retain the enclosing
projection so position-dependent scalar descendants are not flattened. Shape
planning and reference-kind classification likewise treat every reducer as
scalar-valued.

The cacheability walk is iterative and budgeted. A nested statistical reducer
propagates the complete-reference context through its sequence arguments, so a
fixed reference is eligible only when its result is coordinate-independent.
`MUNIT` is handled separately: its scalar size argument does not receive that
complete-reference context, keeping a multi-cell size expression out of the
demand cache and preserving position-sensitive projection. The demand cache
stores only scalar values and formula-error payloads; text, references, arrays,
and typed evaluator failures are not cached.

## Validation evidence

The focused resource suite passed 12/12 tests, including cell and area limits,
work and storage refusal, cancellation, source-version change, unsupported
provider values, lazy unselected branches, typed `IFERROR` refusal, result
ownership, and bounded nested projection. The frozen handoff also records 12
semantic tests, 9 independent numeric-oracle tests, 2 native-profile tests,
and independent native reproduction of 54 observations. The isolated package
gate passed 1,348 tests; all-target Clippy, rustdoc, package format, and the
batch rustfmt gate passed as well. No timing, allocation, RSS, or throughput
claim is made by this review.
