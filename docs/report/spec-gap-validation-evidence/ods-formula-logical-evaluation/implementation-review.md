# Logical evaluator implementation review

Status: closed for the optimized candidate, subject to refreshing the digest if
`evaluation.rs` changes. This is an independent correctness and resource
review of the scalar implementation of the nine OpenFormula logical functions;
it is not a claim of full OpenFormula, reference, array, or recalculation
support.

## Reviewed inputs

* Base implementation: commit `3f907e9e1abcde99b2ce94e88c1122bb79f77a8d`.
* Candidate source:
  `crates/litchi-ods/src/codec/formula/evaluation.rs`, SHA-256
  `cb47a3c6a6a0169f225d0e4b979fe5ec9ebf629bc1a5fdef86dca18f9c3c8efa`.
  The pre-optimization comparison source in
  `performance/before-dispatch/evaluation.rs` is the reviewed
  `3239adc48f7877d4666a912974a5b41ce963559d87376a631af19055fcd4beb3`
  snapshot.
* Focused test source:
  `crates/litchi-ods/tests/ods_formula_logical_evaluation.rs`, SHA-256
  `bea3a987b1a5d922d6ab83e317b8ee997ae3a187685d246b2b7ef3735bb89fad`.
* Normative artifact:
  `3rdparty/specs/OpenDocument-v1.4-os.zip`, SHA-256
  `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`;
  entry `part4-formula/OpenDocument-v1.4-os-part4-formula.html`, SHA-256
  `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.
* Primary anchors: Part 4 §§3.2.3, 5.6, 6.1, 6.2, 6.3.5, 6.3.7,
  6.3.8, 6.3.12, and 6.15.1–6.15.10. Resource and failure boundaries were
  checked against accepted ADRs 0001, 0005, 0006, and 0008.

No Cargo command or production edit was performed for this review.

## Semantic disposition

The implementation has an explicit frame for each lazy control-flow step.
`IF` evaluates its condition and then only the selected branch. Its seven
one-, two-, and three-slot forms use the specified Logical and numeric-zero
 defaults. `IFERROR` and `IFNA` evaluate their first argument once and push the
alternative only when the relevant formula Error is observed. An evaluator
failure such as `Unsupported`, `Cancelled`, `ResourceLimit`, or `Allocation`
never becomes a formula value and is never caught by either handler.

`AND`, `OR`, `XOR`, and `NOT` schedule every supplied argument before applying
the function. A decisive Boolean therefore does not hide a later formula
Error, unsupported reference, cancellation, or limit. `AND` and `OR` use the
`NumberSequenceList` conversion path; scalar numeric Text consequently follows
conversion to Number. `NOT`, `XOR`, and an `IF` condition use direct Logical
conversion, whose deterministic scalar profile rejects Text with `#VALUE!`.
Aggregate application preserves the leftmost source-order formula Error while
still consuming the complete eager value stack. `TRUE` and `FALSE` produce a
separate Logical value at arity zero.

### Malformed arity

The candidate deliberately handles arity classes in two ways:

* `TRUE`, `FALSE`, `AND`, `OR`, `XOR`, and `NOT` use the general eager path,
  including when their arity is invalid. Present children are evaluated and a
  formula Error is produced after their values are consumed; evaluator
  failures remain failures. Zero-argument `AND` and `OR` return the profile's
  permitted formula Error choice.
* `IF` outside one through three children, and `IFERROR`/`IFNA` outside two,
  return `#VALUE!` immediately.

The second rule is an intentional bounded profile decision. Part 4 §3.2.3
requires eager argument computation unless a function specification gives an
exception; §§6.15.4–6.15.6 define the argument positions and lazy behavior
only for the valid `IF` and handler forms. A call outside those signatures has
no defined argument mapping or branch selection, while §6.2 permits an
implementation to reject a constraint violation or document an extension.
Applying the valid-arity lazy algorithm to an extra `IF` argument, or inventing
a missing handler default, would silently define an extension. The immediate
formula Error is therefore consistent with this scalar profile. If a future
compatibility profile accepts extended arities, it must specify and test that
profile separately.

An empty parameter slot is represented by `Missing`. When a required slot is
actually visited, `visit_argument` turns it into formula `#VALUE!`; it is never
fabricated as Empty, zero, or FALSE. The optional `IF` slots are handled by
`push_if_branch` and retain their normative defaults. For an exact two-argument
handler call, the alternative remains lazy: `IFERROR(1;)` returns `1` because
the empty alternative is not selected, whereas `IFERROR(#N/A;)` visits that
slot and returns `#VALUE!`; `IFNA` applies the same rule only for `#N/A`. This
is the direct consequence of the handler equivalence to lazy `IF` and the
profile's selected-branch rule, rather than an optional default for the
handler. The focused tests now cover the complementary cases: `IFERROR(;2)`
catches the visited missing first operand, `IFNA(;2)` returns `#VALUE!`, and
`IFERROR(1;)`/`IFNA(1;)` leave an unselected missing alternative untouched.

## Optimized dispatch comparison

The optimized diff removes the two dispatch-only `Frame` variants and reuses
the existing `Apply(Node)` frame. `visit_function` schedules the same `Apply`
frame at the same point where the prior candidate scheduled `DispatchIf` or
`DispatchIfError`; `apply_function` selects `dispatch_if` or
`dispatch_if_error` before the unchanged eager-function branches. The popped
frame count, value-stack order, selected-branch timing, and all dispatch
checks are therefore unchanged. The frame enum still stores a `Node` payload,
and the variant reduction cannot increase its storage size.

The paired logical harness receipts contain 135 matching rows (45 cases over
parse, evaluate, and parse-evaluate phases). Across every row, the semantic
checksum, success/refusal/failure status, allocation and deallocation counts,
requested/released bytes, live and peak reservation values, and retained
output reservation values are identical between
`performance/before-dispatch/candidate/logical/raw.csv` and
`performance/candidate/logical/raw.csv`. Only timing and process RSS vary.

## Resource and failure disposition

The evaluator remains iterative. The three retained `Frame` variants all
carry the same `Node` payload as the prior frame variants; the two dispatch
variants were removed, so the frame representation does not introduce a
larger per-frame allocation. Lazy dispatch adds only bounded stack entries.
Eager scheduling charges the child count before traversal and checks
cancellation while pushing children; each frame, aggregate loop, and
malformed-arity drain checks the caller context at bounded intervals.

Frame and value vector growth goes through `ensure_capacity`: it checks the
entry limit, checks cancellation, reserves the complete new capacity from the
child Memory budget before `try_reserve_exact`, and restores the old
reservation if allocation fails. Successful growth drops the old reservation
only after the new allocation succeeds. Owned text uses the same
reserve-before-allocation pattern, and the result retains its reservation until
the public value is dropped. Any evaluator error drops frames, values, and
reservations through ordinary Rust ownership.

Logical dispatch itself allocates no argument arrays or cloned values. It
moves `WorkingValue` entries, skips unselected branches, and keeps all source
and cache state immutable. No ambient I/O, resolver, cache publication, or
workbook mutation is introduced. The finite numeric parser and allocator calls
remain indivisible cancellation boundaries documented by the earlier scalar
review; no new logical-specific unbounded boundary was found.

## Conclusion

No correctness or resource blocker was found in the reviewed optimized
`cb47a3c6a6a0169f225d0e4b979fe5ec9ebf629bc1a5fdef86dca18f9c3c8efa`
implementation. The final gate reports 894 tests across 53 targets, two
doctests, Clippy, rustdoc, and formatting passing for this source/test vector.
