# Value API integration decisions

This is an integration checklist following the source-only
[public API review](value-api-review.md). It records remaining work, not API
acceptance or a performance result. The value evaluator and worksheet adapter
are still under development.

## Typed cell refusals

Use `UnsupportedKind::CellValue` when a resolver returns
`CellRead::Unsupported`. A supplied resolver can resolve a cell whose type is
outside the evaluator's supported profile; reporting missing reference support
misidentifies that failure. Keep this operational refusal distinct from a
formula error, including through IFERROR/IFNA. Integration coverage must include
formula-bearing cached values, Date, Time, and Unknown cells.

## Execution policy

Index preparation and evaluation are separate operations. A reusable worksheet
index may retain its construction reservation, charged until drop. Each lookup
during evaluation must receive and charge the evaluation caller's execution
context, including metadata and source-version probes. Pass that context through
the public resolver operations; comparing numeric limits cannot establish
budget or cancellation identity.

Callers can share a hierarchical budget and cancellation token across preparation
and evaluation when they need one aggregate policy. Tests must also construct
the index under an independent preparation policy, then prove that a constrained
or cancelled evaluation policy governs subsequent lookup work. Retaining the
construction policy for all future lookups is not the final contract.

## Remaining value contracts

- Define logical reference geometry admission separately and explicitly from
  cumulative actual provider reads. Preserve both bounds and test bare matrix
  references, scalar projection, repeated reads, and reference lists.
- Require coherent ordered sheet names, extents, and values from every resolver.
  `source_version = None` does not permit inconsistent or mutable observations;
  it means the provider guarantees coherence without a runtime identity fence.
- Give independently allocated equal array and reference-list views structural
  Rust equality, consistent with the existing reference view. Inspection should
  remain allocation-free; document its linear cost. Formula comparison rules
  remain separate from Rust equality.
- Retain borrowed evaluation as the default. Add an explicit budgeted owned
  conversion and a defined owned reference representation for persistent
  results. Verify lifetime independence, exact text preservation, bounded
  allocation, cancellation, and reservation release. Borrowed inspection alone
  does not complete persistent-result ergonomics.

The deferred branch planner must first produce a coherent regression-tested
snapshot. Coordinate these API changes across the evaluator, worksheet index,
tests, and candidate profiling harness afterward. Preserve historical baseline
sources and captures; changed harness bytes require new provenance and an
explicit comparability review.


## Revision 16 implementation evidence

The contracts above are implemented in revision 16; the earlier checklist remains
as the rationale. [The full gate receipt](integration-revision-16.json) binds the
source and passing runtime/documentation checks.

| Contract | Current evidence |
| --- | --- |
| Typed unsupported cell values | `UnsupportedKind::CellValue`; worksheet resolver integration covers inert formula caches and unsupported cell types. |
| Caller execution context | Resolver methods accept the evaluation context; worksheet tests cover independent preparation cancellation and constrained/cancelled lookup. |
| Geometry admission and actual reads | Limit documentation distinguishes both bounds; value tests cover bare Matrix admission and Scalar projection without excess provider reads. |
| Resolver coherence | Public resolver documentation requires coherent sheet ordering, extents and values, including when identity fencing is unavailable. |
| Structural borrowed equality | Independent array/list equality tests cover shape, contents, ordering and duplicates. |
| Persistent ownership | `Evaluated::to_owned` and opaque `OwnedEvaluated`; seven integration tests cover source lifetime independence, exact text/reference metadata, limits, cancellation and reservation release. |

The [owned source review](value-owned-review.md) clears the allocation and
cancellation boundaries. Release profiling is still under review. In particular,
the ownership harness exposed quadratic admission work in a 4,096-entry union
expression during preparation. The isolated union-count fix resolves that failure:
[all 54 debug ownership lanes now pass](reference-union-count-fix.json) under
unchanged limits. Release ownership measurements remain pending. These API implementation checks do not
constitute complete performance or specification acceptance.
