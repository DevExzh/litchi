# Independent review: ODF 1.4 statistical reducers

## Scope and current disposition

This review covers the nine-function statistical-reducer batch defined by
[`contract.md`](contract.md): `COUNT`, `COUNTA`, `COUNTBLANK`, `AVERAGE`,
`AVERAGEA`, `MIN`, `MAX`, `MINA`, and `MAXA`. The contract is based on the
repository-local OpenDocument 1.4 Part 4 archive and records the repository's
finite binary64, empty-text, zero-argument, signed-zero, and resource choices.

The frozen implementation is accepted. All seven isolated gates pass with
1,348 ODS tests and zero failures or ignored tests. The independent resource
review accepts the frozen source. Root independently verified 576 exact oracle
observations, 54 native observations, and all 3,390 performance samples. This
acceptance covers the documented evaluator profile; it does not establish
whole-workbook recalculation or the remaining statistical function families.

## Normative identity

The cited source identity is recorded in the contract:

* `3rdparty/specs/OpenDocument-v1.4-os.zip`, SHA-256
  `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`;
* member `part4-formula/OpenDocument-v1.4-os-part4-formula.html`, SHA-256
  `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.

The relevant normative boundaries are the sequence and Any pseudotypes,
common error propagation, the three count sections, and the six statistical
reducer sections. The existing aggregate contract supplies the exact
binary64 sum state and hierarchical resource rules; it is not a substitute
for tests of this family's distinct typed admission behavior.

## Semantic review checklist

The implementation review must establish all of the following from source and
tests:

1. `COUNT`, `MIN`, and `MAX` use `NumberSequenceList`; `AVERAGE` uses
   `NumberSequence` and rejects an explicit ReferenceList; `COUNTA`,
   `AVERAGEA`, `MINA`, and `MAXA` use `Any` and admit ordered ReferenceLists;
   `COUNTBLANK` accepts only ReferenceList/Reference values.
2. Reference sequence traversal is list occurrence, sheet-plane, then
   row-major order. Duplicate and overlapping references are not deduplicated.
   A logical 3-D Reference remains a single Reference and is not implicitly
   intersected or flattened into a caller-shaped array.
3. Referenced Text, Empty, and distinguished Logical values are omitted by
   NumberSequence reducers. Scalar Text and Logical values use the finite
   Number conversion profile. Inline arrays use the documented row-major value
   extension and do not authorize a ReferenceList-to-Array conversion.
4. `COUNT` ignores formula Errors in references and direct Error/conversion
   Error values (while typed evaluator failures remain hard failures);
   `COUNTA` counts every non-empty value including Errors and empty Text;
   `COUNTBLANK` treats Errors and non-empty values as non-blank and, by the
   selected host-defined choice, treats empty Text as blank.
5. Each supplied missing argument is normalized to one formula `#VALUE!` in
   both scalar and value evaluation: `COUNT` suppresses each, `COUNTA` counts
   each, and the other reducers propagate the first. The parser's `COUNT(;)`
   and `COUNTA(;)` forms therefore contain two missing arguments and return
   `0` and `2`, respectively. Zero arguments remain distinct from supplied
   missing slots. A reducer returns one scalar; any matrix broadcast comes
   from the enclosing array operator.
6. `AVERAGE` and `AVERAGEA` use `#DIV/0!` for an empty included set; all four
   extrema functions use numeric `+0` for an empty included set and
   canonicalize selected zero results to `+0`.
7. `A`-family reducers include Text as zero and Logical as `0`/`1`, omit Empty
   cells, and preserve formula Errors. Complex values are rejected by the
   bounded real reducer profile rather than silently projected.
8. Formula Errors, generated formula errors, and typed evaluator failures
   follow the precedence and publication rules in the contract. A typed
   unsupported/read/cancellation/source-version/resource failure must not be
   converted into a formula Error or hidden by `IFERROR`.
9. Reducers nested in projected matrix branches consume complete references;
   scalar descendants that vary by output coordinate are not demand-cacheable.
   Cache classification is iterative and budgeted.

Any mismatch in this list is a semantic blocker, even if a native spreadsheet
cache happens to agree with the implementation.

## Numeric and memory review checklist

The source review must confirm that:

* average reductions reuse the checked exact sum state and divide once by a
  checked count, preserving cancellation, minimum-subnormal results, and
  negative underflow to `-0`;
* extrema compare finite binary64 values without arithmetic accumulation and
  publish a deterministic `+0` for either signed zero;
* all references are streamed without a full cell-vector materialization;
* every physical read and reducer operation is charged to the caller's work
  budget, metadata and text limits are checked before allocation, and
  cancellation/source fences are checked before publication; and
* no production path recalculates formula cells, refreshes cached values, or
  clones borrowed resolver text per cell.

The review cannot accept an implementation that claims boundedness only from a
test's small fixture. It must inspect the allocation and drop order in the
actual reducer path and verify low-ceiling refusal tests.

## Required evidence before acceptance

The final report should include the focused scalar/value semantic test counts,
the exact oracle row count and functions, any native observations and explicit
cache exclusions, all seven gate outcomes, and a separate performance receipt
with matched controls. The evidence should include mixed typed cells,
ReferenceLists, 3-D references, empty text, Errors, empty averages, extrema
signed zero, exact cancellation, negative average underflow, projected cache
cases, unsupported cells, work/storage ceilings, cancellation, and source
version changes.

Those receipts are now retained and independently verified; the final
disposition below accepts this bounded statistical reducer profile.

## Final receipt disposition

The source hashes in `gates/freeze.json` match the implementation and final
compiled closure. Semantic, resource, numerical, native, strict warning,
formatting, boundary, and whitespace checks pass. No debug output remains.

Root recomputed all performance medians and resolver-read bounds from the raw
390 baseline and 3,000 candidate samples. Allocation counts, requested bytes,
and peak live bytes are unchanged on matched controls. Unrounded median time
shifts are −5.46% to +1.90%; RSS shifts are −4.09% to +3.92%. No positive
matched metric exceeds the 5% review threshold. The profile is bounded to the
recorded fixture, machine, compiler, and evaluator scenarios; it supplies no
whole-workbook or cross-platform performance claim.
