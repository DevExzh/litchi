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
`e9d8dbac4eb9d9fe964693ccb801e74a031f7c223b3595e5003de03e36f02288` and it
contains 52 selected inputs. The gate's source-before and source-after
manifests match all 52 frozen inputs; the live shared root matches every
selected source path except the intentionally separate ambient `Cargo.lock`.
The principal implementation, contract, and focused-test hashes are:

| Frozen file | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `b378c4c962814dbe7602ab315569603563bfbe42c62042af14a307adc1c66a48` |
| `crates/litchi-ods/src/codec/formula/evaluation/reference_metadata.rs` | `0bf2802eb9cdde75bb45fba4cd5dbb78234bdd6f73e8139788ad9f9591912eb2` |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `7adbcc55dc985f39560379748b00bf43b9413d0c45f2cca0b280cb045a420dc4` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/reference_metadata.rs` | `bfeb375b63875b79e68cd4b6dd38b46765f083e47b83b79ef06da092a6af43be` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/references.rs` | `87f704ea35549858f357083b3675a307ef777241bb81b458027735742cccf03d` |
| `crates/litchi-ods/tests/ods_formula_reference_metadata_evaluation.rs` | `c0e10746ddb05767642875c8ea72544d86ff382fc6f56bb9c100981b21bb7fa4` |
| `crates/litchi-ods/tests/ods_formula_reference_metadata_limits.rs` | `ca88559e2d6bf6629eaffb10003cfc9c1de4f6f1a1fa088909e2daefb040e622` |
| `crates/litchi-ods/tests/ods_formula_reference_metadata_oracle.rs` | `7b4fee554d92305fd53343383c4959e84c958103b30a9a90020c5696ba1b4247` |
| `docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata/contract.md` | `87f20c419139bf721221d06c8714edd900ae7a57cbedecf070878e4f6cf64d8a` |
| `docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata/reference_metadata_oracle.py` | `00b3e1b5db99bfa1a8d58fc2866e4a586cbbab843df36214e7ca06dfca7ddde9` |
| `docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata/reference-metadata-goldens.json` | `b75548f94c5f1eff963b2ae0d520a8b3fe4ec4a1319c47f00a6c35e005171f53` |

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

The frozen focused handoff supplied:

| Target | Result |
| --- | --- |
| `ods_formula_reference_metadata_evaluation` | 16/16 |
| `ods_formula_reference_metadata_limits` | 7/7 |
| `ods_formula_reference_metadata_oracle` | 1/1 test, 84 observations |

The focused cases cover strict one-entry and multi-entry list refusal, 3-D
record/plane geometry, fixed current-position metadata, projected computed
SHEET arrays, direct and selected source leaves, scalar error-handler fallback,
matrix source-order isolation, zero cell reads, typed metadata failures,
cancellation/resource/source fences, output ownership, and demand-cache
classification. Strict all-target Clippy also passed at handoff. Native
spreadsheet observations are retained by the bundle owner. The isolated gate
receipt reports all seven required checks passing, stable source-before/source-
after manifests, and 1,647 package tests with no failures or ignored tests.
This review makes no timing, allocation, RSS, or throughput claim.
