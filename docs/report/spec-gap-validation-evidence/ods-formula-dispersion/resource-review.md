# ODS dispersion reducer resource and cache review

This review covers the resolver-aware implementation of `VAR`, `VARA`,
`VARP`, `VARPA`, `STDEV`, `STDEVA`, `STDEVP`, and `STDEVPA`. It applies
ADR 0005's finite work, storage, source, and cancellation rules and ADR
0006's typed-failure and publication rules to the frozen
[dispersion contract](contract.md). The review changed no production code or
tests.

The frozen contract SHA-256 is
`5c4b881fcddee6bd5a1d4a143d74fbfb794c4d58726e3922ae492c963a516852`.

The disposition is **pass** for the resource and demand-cache boundaries. The
review found no remaining resource blocker in the frozen implementation.

## Frozen source identity

The baseline is commit `55e147bfa0676ce6ecdc609efc682b98568b8a5f`. The source
freeze is recorded in
[`gates/freeze.json`](gates/freeze.json); each selected file was re-read after
the freeze and matched that manifest:

| Frozen file | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `fb9ebd5e684fa2cde84d5fce5960e77184571bffc47ec3ef8bcbe73a8dc40ed4` |
| `crates/litchi-ods/src/codec/formula/evaluation/statistical.rs` | `1045c81b8d7780b6d4edd5073f66240701b2f7e6c4a2ad3b894e326f9eb0328a` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/statistical.rs` | `6d6b1d1432d4a9367c4c29e1a672aed0e92368fb149c8615df8bd03292fcdbf3` |
| `crates/litchi-ods/tests/ods_formula_dispersion_evaluation.rs` | `ab89fb34f4053444320b00d2a3479bee50cf7c7171fe9c211b442edc1ad34693` |
| `crates/litchi-ods/tests/ods_formula_dispersion_limits.rs` | `1d20d12e1c99d2e926f9993880bc96035411bbd3d48d4544f93a1739208d9244` |
| `crates/litchi-ods/tests/ods_formula_dispersion_native.rs` | `223f05af0b81613059fc49b0f5b20e3e6e2ed82f566eb9e78b263ed0b9ed2fc0` |
| `crates/litchi-ods/tests/ods_formula_dispersion_oracle.rs` | `f42a9a96b5bb26b82b623b4c26b8e6fa22c4459c59a280ddadd5701cdc687825` |
| `docs/report/spec-gap-validation-evidence/ods-formula-dispersion/numeric-goldens.json` | `0a5a52c6f38e045f0221038d4563a845b4730845579068e7a042532469df9e8b` |
| `docs/report/spec-gap-validation-evidence/ods-formula-dispersion/native/cached-results.json` | `88c785f9b36e98a8b7e88f0bbfc76ad99b845d1592070ea1c32feea7e46f02f5` |
| `docs/report/spec-gap-validation-evidence/ods-formula-dispersion/native/provenance.json` | `d1f1e0317ae44a7b03178d0a0b4fe19bac6d9561f439123eee53522d2cf8ef35` |

The isolated gate checkout used `Cargo.lock` SHA-256
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`, as
recorded by the same freeze manifest. The ambient workspace lockfile is not
the gate dependency input.

## Resource and ownership checks

* **Geometry and storage admission:** `RuntimeAreaSet` validates rectangle
  cell counts and cumulative reference cells with checked arithmetic before a
  reducer scan. Retained area and reference metadata, AST frames, runtime
  arrays, shape/reference planners, and demand-cache entries use fallible
  `ensure_capacity` reservations under the evaluator storage budget. The
  dispersion reducer keeps only fixed `StatisticalState` and
  `NumericAggregate`/variance state; it never materializes a reference-cell
  vector for a scalar result.
* **Work, reads, and cancellation:** `scan_reference` visits retained areas
  in area, sheet, row, and column order. Every physical cell charges work
  before `read_reference_cell`. That read path checks the cumulative reference
  limit and cancellation before the resolver call, checks cancellation again
  afterward, and increments the successful-read count only after the provider
  returns. Array cells and direct argument inspection use the same bounded
  work accounting.
* **Typed failures and formula errors:** Each dispersion reducer retains the
  first formula error in source order and continues the admitted scan. A
  later resolver `Unsupported`/resource/cancellation/source failure remains
  an `EvaluationFailure`, so it supersedes the retained formula error and is
  never converted to a formula error, cached, or caught by `IFERROR`.
* **Source and publication fences:** The value evaluator checks cancellation
  and the resolver source version before evaluation, checks the source again
  after all reads, and checks cancellation once more before publishing the
  retained scalar result.
* **Borrowed text:** `read_to_element` applies the per-value text limit and
  creates a borrowed `TextValue`. A-variant text-to-zero conversion charges
  the borrowed text bytes without cloning one provider string per cell.
* **Bounded state and drop order:** The fixed variance accumulator has no
  per-cell allocation. Runtime area/record and array buffers are paired with
  their reservations so retained elements are dropped before the reservation
  token on success and typed-failure paths.
* **Zero-read shape refusal:** `VAR`, `VARP`, and `STDEVP` reject an explicit
  `ReferenceList` before `scan_reference`, yielding formula `#VALUE!` without
  a resolver cell read. `STDEV` and all four A variants admit ordered lists.
  Direct scalar/array shape and resource refusals likewise occur without
  resolver reads. A computed reference expression may read while it is being
  evaluated before this descriptor gate; the zero-read guarantee applies to
  the already-retained refused descriptor.

## Demand-cache checks

The statistical family is included in projected-branch sequence classification
before argument scheduling and in the apply-time cache get/put path. Literal
arrays use `VisitMatrixArgument`; retained references keep complete area/list
descriptors rather than being implicitly intersected at the output coordinate.
Shape planning and reference-kind classification mark each reducer as a
scalar result. Computed reducers stay conservative when a projected scalar
descendant can depend on the output position.

The iterative cacheability walk propagates complete-reference context through
statistical reducer arguments and nested conditional criteria. This permits a
complete fixed-reference reducer to be reused across projected outputs while
keeping ordinary position-dependent arithmetic out of the cache. `MUNIT` is
handled separately: its scalar size argument is excluded from full-argument
propagation, so a nested MUNIT criterion is recomputed at each required
position. Demand-cache entries contain only scalar value/error payloads and
are evaluator-local to the source/context-fenced evaluation; typed evaluator
failures are never cached.

## Validation evidence

The frozen focused suite passed 30/30 tests:

* 9 semantic evaluator tests;
* 11 resource and typed-failure limit tests;
* 8 independent-oracle tests over 512 observations and 56 fixtures; and
* 2 native-profile tests covering 48 observations from eight pinned FODS
  inputs.

All seven isolated gates pass against unchanged source: 1,378 tests with
zero failures or ignored tests, all-target Clippy with warnings denied, rustdoc
with warnings denied, package formatting, selected-file formatting, crate
boundaries, and diff whitespace checks. Root independently verified the
receipts in `integration-verification.json`. This review makes no timing,
allocation, RSS, or throughput claim.
