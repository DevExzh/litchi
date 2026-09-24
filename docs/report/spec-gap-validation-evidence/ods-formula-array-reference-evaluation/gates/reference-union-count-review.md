# Reference union cached-count review

This is a source-only review of the `RuntimeAreaSet::cell_count` union
optimization. No Cargo build or test run is claimed. The review is scoped to
reference-area construction, intersection materialization, union admission,
checked limits, cancellation checks, and reservation cleanup; it is not a
release or general value-VM acceptance review.

## Reviewed source

| file | SHA-256 |
| --- | --- |
| [`evaluation/value.rs`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value.rs) | `9569c8aa98a1d2f9e7a14bc9057d1b5cba33d4adb6257534afbe21a82f1c760b` |
| [`evaluation/value/references.rs`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value/references.rs) | `680a3fd9c3b80951408f3210f5c5373dcea7dd00c1b561b0902043438b79d907` |

## Disposition

The cached-count invariant is cleared for the reviewed paths. Every current
physical-area insertion either goes through `push`/`push_raw` or is a
constructor initialization. `RuntimeAreaSet::empty` and the reference-error
coercion initialize the count to zero; direct and derived sets begin at zero;
range construction, direct references, and intersection output update it via
`push_raw`.

`push_raw` checks the rectangle cell count and checked-adds it to the retained
count before capacity admission. The count is assigned only after the vector
capacity is admitted and the area is pushed. A failed limit, cancellation, or
allocation path therefore cannot publish a count for an unretained area.
Intersection preflight sums the same physical intersection rectangles that
materialization inserts, including duplicate list entries. Public area merging
only changes metadata representation and does not alter the physical count.

`append_set` drains and adopts the right-hand physical areas through
`push_raw`, then moves records after destination record capacity is admitted.
The source and destination reservations remain owned by the corresponding
live buffers and are released on normal completion or when a partial local
result is dropped. Duplicate and overlapping areas remain counted separately,
as required by the existing reference-list contract.

## Work-accounting boundary

`union_ranges` now admits the planned cell total from the two cached counts,
so it no longer rescans the retained left and right area vectors. The removed
rescans also no longer incur their old per-area `Work` charges. This is the
accepted accounting contract: the work budget measures work actually
performed, and retaining a cached total is constant-time work. The right-hand
adoption loop still charges each area it processes, and ordinary VM/operator
work remains charged by its existing paths.

This can change which typed resource refusal wins when a cell limit and a local
work limit would both have failed under the old rescan implementation. There is
no promise of the old refusal ordering; both remain typed atomic refusals. The
optimization must not reintroduce phantom charges solely to preserve that old
ordering.

No stale-count, unchecked accumulation, changed geometry admission, or
reservation leak was found in the reviewed current mutation paths. Cancellation
checks remain on right-area adoption and on capacity growth; the optimization
does not claim a general value-VM or performance acceptance result.
