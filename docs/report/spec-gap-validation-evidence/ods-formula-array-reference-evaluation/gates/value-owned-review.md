# Owned value conversion review

This source review is **resolved for the previously identified ownership and
cancellation boundaries**. It is a source review only; no build or test was
run in this pass, and it does not constitute full value-VM acceptance.

Reviewed source hashes:

| file | SHA-256 |
| --- | --- |
| [`value/owned.rs`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value/owned.rs) | `c9392c7e6b93a58d8521711da7adec10938bfd895d94244c3417c97d44a670c2` |
| [`value.rs`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value.rs) | `bd7605b6f000cc470787b881fd1b661468d826a4d4a2c41bf33ed6a6fb834015` |

## Confirmed properties

* Conversion remains a measure-then-copy operation. The first pass accounts
  aggregate `Memory` and `Work`, exact vector capacities, text bytes, arrays,
  areas, and every nested reference component before publication. The single
  reservation is attached to the result, and `OwnedValue` is declared before
  that reservation, so owned storage drops before its charge is released.
* Reference ownership is complete and fallible. Source IRIs, sheet names,
  subtable chains, columns, vectors, areas, absolute/quoted markers, duplicate
  records, and ordering are copied through checked `try_reserve_exact` paths.
  Partial copies and allocation failures unwind without publishing a result.
* Evaluated text is checked against `max_text_bytes`; reference lexical
  strings are included in aggregate owned storage. UTF-8-safe chunked copying
  preserves the exact text value and gives cancellation checkpoints for large
  strings.
* The former cancellation admission gap is closed. `reserve_storage` checks
  the caller before child-budget creation and immediately before the direct
  child-budget reservation. `copy_string` checks immediately before its
  allocation, and `fallible_vec` accepts the execution context and checks
  before every nonempty allocation. The added module regression covers
  cancellation after measurement and verifies that storage remains uncharged.
* The final copy fence still checks cancellation after all owned data has been
  formed. A failure in either pass therefore leaves the source result intact,
  drops partial owned values, and releases any reservation already acquired.
* Owned array and reference-list views compare structurally, preserving shape,
  cell values, reference metadata, and record order. The parent borrowed
  `ArrayView` and `ReferenceListView` now use the same structural semantics and
  document that equality is linear inspection rather than OpenFormula
  coercion.
* `reserved_output_bytes` now explicitly documents that it aliases the
  retained `Resource::Memory` charge; it does not imply a separate output
  resource reservation.

## Scope and remaining follow-up

The public owned-conversion integration checkpoint reported seven passing
tests covering lifetime independence, Unicode and escaped text, 3-D and
duplicate references, lexical markers, structural equality, resource limits,
cancellation, and reservation release. This review did not rerun that suite.

The owned result intentionally exposes an opaque owner with borrowed views;
direct extraction of independently owned strings or collections remains an
ergonomic follow-up. The conversion review does not assess the value VM's
reference evaluation, worksheet resolver integration, or full production
acceptance.
