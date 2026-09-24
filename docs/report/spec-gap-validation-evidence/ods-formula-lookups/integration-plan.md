# ODS lookup and reference-function implementation

Status: active implementation batch; no support or passing-gate claim yet.
Baseline and preserved unrelated changes are recorded in `baseline.json`.

The nine-function batch comprises ADDRESS, CHOOSE, HLOOKUP, INDEX, INDIRECT,
LOOKUP, MATCH, OFFSET, and VLOOKUP. GETPIVOTDATA and MULTIPLE.OPERATIONS remain
tracked gaps requiring pivot calculation and dependency recalculation services.
The broader audit objective remains active.

## Ownership

- Shared private catalog, scalar conversion/comparison and address syntax:
  `evaluation/lookup.rs` and its children.
- Runtime scheduling, lazy CHOOSE, descriptor selection, geometry and dynamic
  references: `evaluation/value/lookup.rs`, its selection children, and VM hooks.
- Search kernels: `evaluation/value/lookup/search.rs`.
- Independent semantic/resource tests: `ods_formula_lookup_evaluation.rs` and
  `ods_formula_lookup_limits.rs`.
- Independent oracle/native fixtures and performance harness: dedicated owners
  under this evidence directory.
- Independent semantic and resource/cache reviews after source stabilization.
- Root owns integration gates, evidence acceptance, source freeze and commits.

## Implementation invariants

Keep references as descriptors until a consumer demands values. INDEX and OFFSET
must compose with metadata, reducers, reference operators and selected lazy
branches without materializing unrelated cells. INDIRECT translates temporary
parsed syntax into derived geometry using resolver-owned canonical sheet names;
no temporary text borrow may escape. Source-qualified references stay inert.

CHOOSE evaluates only its index and selected expressions; projected array
selection follows the existing lazy IF resource and source-policy model. Shape
and kind discovery must not catch public provider errors as internal control
flow. A typed resolver failure, cancellation, resource failure or source change
cannot become a formula error or be caught by IFERROR/IFNA.

Exact lookup scans in order and returns the first match. Sorted lookup may use
bounded binary search and must select the last duplicate. Read only needed search
cells and the selected result. Use borrowed text and fixed-size comparison state;
AST arrays and result arrays remain bounded, charged allocations. Scalar versus
ForceArray argument roles are explicit. Cache only proven invariant scalar
payloads and retain position sensitivity through scalar descendants such as MUNIT.

## Acceptance

Freeze the normative contract before implementation expectations are finalized.
Validate all nine functions across scalar, matrix, projected, reference/list/3D,
invalid-input and resource cases. Compare an independent oracle and preserve
native compatibility differences explicitly. Run focused tests before seven
isolated gates: ODS tests, strict all-target Clippy, strict rustdoc, package and
selected-source formatting, boundaries and diff checks. Preserve the ambient
Cargo.lock and run isolated gates with the retained lock identified in baseline.

Capture performance only after source freeze and independent reviews, using
matched controls and lookup scaling cases. Retain all attempts and all review
flags; normalize raw elapsed time before summarizing. Bind source, inputs, raw
samples and reports with hashes, verify retained evidence after cleanup, and
commit only this batch. Do not infer end-to-end workbook speedup from evaluator
measurements.

## Root review probes

The final source review will specifically test runtime-kind discovery for
failing reference constructors (for example ISREF(INDIRECT("bad text"))).
Classifying a function as reference-producing must not bypass its argument
errors or claim a valid reference when construction failed. Dynamic references
inside range/intersection/union must propagate provider failures without retry.

ADDRESS formatting must quote sheet names so INDIRECT can parse them, including
apostrophes, non-ASCII names, names beginning with digits and cell-like names.
R1C1 ADDRESS uses supplied coordinate numbers without subtracting formula
position; R1C1 INDIRECT applies bracketed offsets to the explicit position.

LOOKUP short-result extension is conditional on the selected match falling
outside the existing result vector. A near-edge result reference with an
in-range match remains valid. Arrays are never extended; a longer result vector
is not shortened. Selector formula errors retain their original error value.

Integer conversion must reject values outside the destination integer domain
before casting. In particular, `i64::MAX as f64` rounds to 2^63, so an inclusive
comparison against that floating-point bound would admit an invalid value and
silently saturate. Missing optional slots take defaults; actual Empty values
still undergo the declared conversion.

Coverage acceptance must bind every function requirement and cross-cutting
requirement to concrete retained evidence. A nonempty JSON object, a planned
test name or a planning document is not a passing receipt. The final verifier
must reject missing bindings, altered evidence and paths outside the retained
source/evidence closure.
