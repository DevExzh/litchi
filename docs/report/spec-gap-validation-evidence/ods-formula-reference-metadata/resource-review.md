# ODF reference-metadata resource and cache review

This review covers the eight OpenFormula 1.4 §6.13 functions `AREAS`,
`COLUMN`, `COLUMNS`, `ISREF`, `ROW`, `ROWS`, `SHEET`, and `SHEETS`. It applies
ADR0005's bounded work, storage, cancellation, and source rules and ADR0006's
typed-failure and publication rules to the profile in
[`contract.md`](contract.md). The review made no production or test edits.

The disposition is **PASS** for the resource and demand-cache boundaries. The
source-policy fixes in the frozen handoff preserve scalar `IFERROR`/`IFNA`
fallback metadata context and clear metadata-only policy for matrix value
arrays. I found no remaining resource or cache blocker.

## Frozen source identity

The freeze manifest is [`gates/freeze.json`](gates/freeze.json), based on
commit `049c09cdde3978593149079c4257df047a3fa419`. Its SHA-256 is
`b4a0f0c9665e10f67fafa1b179bfbfdd8be2c65e6ec897bac86db1b2afc9330a` and it
contains 52 selected inputs. This corrected manifest includes the final
`ReferenceShapeOutcome::Refused` separation. The live shared root matches every
selected source path except the intentionally separate ambient `Cargo.lock`.
The principal implementation, contract, and focused-test hashes are:

| Frozen file | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `b378c4c962814dbe7602ab315569603563bfbe42c62042af14a307adc1c66a48` |
| `crates/litchi-ods/src/codec/formula/evaluation/reference_metadata.rs` | `0bf2802eb9cdde75bb45fba4cd5dbb78234bdd6f73e8139788ad9f9591912eb2` |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `5d5ca235bf9d578422e058e0720d364d84bb84a783f6b0ff0d8fbe4899e7b2b2` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/reference_metadata.rs` | `f845d12b72d365c441352e13b170e58e8671f418f473656f6dc7edcd0e5c0d7a` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/references.rs` | `87f704ea35549858f357083b3675a307ef777241bb81b458027735742cccf03d` |
| `crates/litchi-ods/tests/ods_formula_reference_metadata_evaluation.rs` | `c3176d76eec0df3b0a32b3790ce581ddd998e9a41e6e08fd04ad65c2d6831195` |
| `crates/litchi-ods/tests/ods_formula_reference_metadata_limits.rs` | `b4a720a527069c0ffc7e9f2e0c77acce4a3030c9e4a99e73254d7943f3189828` |
| `crates/litchi-ods/tests/ods_formula_reference_metadata_oracle.rs` | `9bf255f51c91d1779bec9f9c5d2d1dd5d210fc39a361ad5ea61dfcc6c74d71e2` |
| `docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata/contract.md` | `87f20c419139bf721221d06c8714edd900ae7a57cbedecf070878e4f6cf64d8a` |
| `docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata/reference_metadata_oracle.py` | `bde7fb9a08251eebdefe5071494bb55d3d9a56fb2421ab0e2a2517159410844d` |
| `docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata/reference-metadata-goldens.json` | `b5c3e55a13be348c426e7df609cbc07e94e033f6ba762ba28ec8e5973702f241` |

The isolated gate checkout uses the frozen `Cargo.lock`, SHA-256
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`. The
ambient root `Cargo.lock` is
`aa945c79965460e74a64063e1c45396eaae7a68e71072730e2bc426cded22f02`; this
intentional workspace mismatch does not change the frozen source identity or
the gate lock used for validation.

## Resource and ownership checks

* **Descriptor admission and bounded storage:** `reference_areas` computes and
  checks rectangle cells, physical sheet span, cumulative `max_reference_cells`,
  and `max_reference_areas` before retaining geometry. `RuntimeAreaSet`, its
  logical records, AST frames, shape masks, and cache scratch vectors use
  checked arithmetic and fallible `ensure_capacity` reservations. A 3-D
  reference retains one logical record while its physical planes remain
  ordered; `AREAS` counts records and `SHEETS` counts planes.
* **Metadata work and cancellation:** Descriptor expansion, reference
  operators, axis output, text-array mapping, cache walks, and shape planning
  charge scalar or cell work through the existing execution budget. The
  evaluator checks cancellation at the established frame, loop, and provider
  boundaries. Resolver metadata failures, cancellation, resource limits, and
  source changes remain typed `EvaluationFailure` values.
* **No cell materialization:** The metadata implementation uses retained
  reference geometry, sheet indices, extents, and ordered sheet names. It does
  not call `read_reference_cell`, `read_cell`, or materialize a reference to
  answer `AREAS`, `COLUMN`, `COLUMNS`, `ISREF`, `ROW`, `ROWS`, or `SHEETS`.
  `SHEET(Text)` performs only the required sheet metadata lookup. Known
  pseudotype, list, source, and geometry-limit refusals therefore make zero
  provider cell reads; descriptor admission may still call
  `sheet_extent`/`sheet_index`/`sheet_name_at`, as permitted by the contract.
* **Output reservation:** `COLUMN` and `ROW` reserve their complete generated
  arrays with `new_element_vec` before filling them. Matrix `SHEET` text output
  follows the same fallible output reservation path. A projected scalar axis
  or text-array element allocates no second function-local result; the outer
  matrix result owns and bounds the published output. Checked shape and count
  arithmetic precede allocation, and retained cells are dropped before their
  reservation tokens on success and typed-failure paths.
* **Borrowed text and conversion work:** `SHEET(Text)` retains parsed or
  resolver-provided text without turning a reference into a cell value. Text
  mapping charges the text conversion/lookup work and per-value bytes through
  the scalar budget. `read_to_element` remains the borrowed text boundary for
  the shared resolver path.
* **Formula errors and typed failures:** Formula Errors remain values. `ISREF`
  classifies them as `FALSE`; the other metadata functions propagate an
  admitted Error or produce their documented formula `#VALUE!` pseudotype
  result. Source arithmetic, unsupported external geometry, provider failure,
  cancellation, resource exhaustion, and source-version changes are never
  converted into formula Errors, cached, or caught by `IFERROR`.
* **Source and publication fences:** A source-qualified leaf is represented by
  a bounded `SourceReference` marker. `ISREF` can inspect it without provider
  access, while `SHEET`/`SHEETS` reject it at their function boundary with
  formula `#VALUE!`. Scalar `IFERROR`/`IFNA` restores the saved source policy
  before selecting a fallback. Matrix `IF`/`IFERROR` clears that policy before
  each value cell, so a selected external leaf requires the typed external
  provider and cannot leak descriptor state from a prior cell. Full evaluation
  source-version and final-cancellation fences remain around publication.

## Demand-cache checks

Metadata functions participate in the projected-branch cache lookup before
argument scheduling and in the apply-time cache get/put path. The scheduler
uses complete argument contexts for reference geometry, `ISREF`, `SHEET`, and
`SHEETS`; list/reference descriptors are classified before scalar projection.
The cacheability walk is iterative and budgeted, and only scalar values and
formula-error payloads enter the demand cache.

No-argument `COLUMN()`, `ROW()`, and `SHEET()` use the fixed formula origin
across projected output coordinates; they are conservatively left uncached.
Axis arrays are not cached as scalar payloads. Stable literal text, local
descriptors, `SHEETS()`, and invariant complete descriptor operations may be
reused; text-array and computed metadata expressions remain conservative when
their output can depend on the projected coordinate. Direct `ISREF` may cache
the invariant runtime kind, including a source marker or ReferenceList,
without any cell read.

Nested metadata calls inside conditional criteria are conservatively excluded
from the outer criterion cache. Sequence reducers propagate complete reference
arguments only in their reducer context; existing `MUNIT` scalar arguments
remain position-sensitive and excluded from complete-reference propagation.
Typed provider/resource/cancellation/source failures are never cache entries,
and cache lookup does not bypass the evaluator's source or cancellation fences.

## Validation evidence

The current frozen focused handoff supplied:

| Target | Result |
| --- | --- |
| `ods_formula_reference_metadata_evaluation` | 19/19 |
| `ods_formula_reference_metadata_limits` | 12/12 |
| `ods_formula_reference_metadata_oracle` | 1/1 test, 91 observations |

The focused cases cover strict one-entry and multi-entry list refusal, 3-D
record/plane geometry, fixed current-position metadata, projected computed
SHEET arrays, direct and selected source leaves, scalar error-handler fallback,
matrix source-order isolation, zero cell reads, typed metadata failures,
cancellation/resource/source fences, output ownership, and demand-cache
classification. The shared-root and isolated frozen package suites passed
1,655 tests with no failures or ignored tests. All seven isolated gates passed,
and source-before/source-after are identical. The final gate receipts are:

| Receipt | SHA-256 |
| --- | --- |
| `gates/results.json` | `40dd5b4d5d215bc51de1030ad37e1c8a854e86732a8ff865a543e6272f42329a` |
| `gates/verification.json` | `bdeb941ccf550e43a341df081dcd54e1ba9e70779744ecb80e785cca71739243` |
| `gates/source-before.json` | `193f42d6618cf2a732a8b54b86ad823d048315fc08464527c4f256429ebb00d3` |
| `gates/source-after.json` | `193f42d6618cf2a732a8b54b86ad823d048315fc08464527c4f256429ebb00d3` |

The individual gate logs record zero exit codes for ODS tests, warning-denied
Clippy, rustdoc, formatting, batch formatting, crate-boundary checks, and
`git diff --check`. Native spreadsheet observations are retained by the bundle
owner. Performance acceptance remains pending its final receipt; this review
makes no timing, allocation, RSS, or throughput claim.
