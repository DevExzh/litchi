# Value planner review, revision 15

This is a source-only review of the deferred shape planner and scalar probes in
[`value.rs`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value.rs).
It is a bounded review of nested demand propagation, cache keys, accounting,
and lazy branch materialization. It does not re-audit the scalar bridge,
reference operator implementation, worksheet resolver, or public owned-value
API.

## Reviewed source

The requested revision identifier was
`d5dff62b4a162c87a8e6c2b39b9c1ec2da02f73c581f033f54f21e99fd6250cd`. That hash
was not present in this checkout. The integration-revision-16 receipt
identified its tested snapshot as `bd7605b6f000cc470787b881fd1b661468d826a4d4a2c41bf33ed6a6fb834015`;
the shared source advanced during this review with an unrelated reference-set
accounting change. The latest observed source file is

```text
crates/litchi-ods/src/codec/formula/evaluation/value.rs
284c3c500bebda56daecd8e4c302b39d983438f1fa0d7096c5b271c7536c1be4
```

The receipt records the surrounding test, Clippy, rustdoc, doctest, and format
commands as successful. Those commands were not rerun for this source-only
review, and passing tests do not cover the demand-growth cases below.

## Findings

### VP15-1: condition-cache keys do not normalize broadcast coordinates

**Blocking for cumulative read limits and general nested shape growth.**

`condition_cache_position` (around lines 2543–2568) keys an entry by the exact
node, demand rows, demand columns, and output index. `evaluate_scalar_at` first
stores the source key `(condition, condition_shape, condition_index)`, while
`deferred_lazy_children` additionally stores an exact effective-demand alias
for every selected output index (around lines 3270–3275 and 3326–3331). The
aliases are useful only for that exact shape.

`matrix_branch_shape` replans each branch until that branch's own demand
converges (around lines 2189–2251), but starts the next branch from the
original condition shape. If the later branch widens the final matrix, the
earlier branch is not replanned at the final shape. Matrix emission then uses
the final projection shape (around lines 2800–2805), so the earlier aliases are
misses even when the output coordinate projects to the same source condition
cell.

A concrete profile case is:

```text
=IF({TRUE();FALSE()};IF(XOR([.A1:.B1]);{1};0);{0|1|2|...|(M-1)})
```

With `A1` true and `B1` an unreadable cell, the first outer column selects the
nested branch. The first branch converges at the initial `1 x 2` demand; the
second branch widens the final result to `M x 2`. For every first-column row,
the nested `XOR` projects the same row-vector source cell `A1`, but its
condition-cache lookup uses `(M, 2, index)` and misses the earlier `(1, 2,
0)` entry. Because range references are not in the coordinate-independent
`demand_cache`, `A1` is read once per row. Setting
`max_reference_cells = M - 1` makes the current evaluator refuse a result
whose only required source cell could have been read once; setting the array
limit to at least `2*M` isolates this from output admission.

The cache lookup should normalize an output coordinate through the condition's
source shape (or the planner must replan every selected branch after the final
shape grows). Under the current fixed written-coordinate profile, this does
not require adding a source-origin key: the same evaluator position and
written endpoint identify the same source cell. The earlier concern about an
origin collision is therefore withdrawn; this finding is specifically about
shape aliases and sibling growth.

The same aliasing also creates avoidable entries. For a scalar condition and a
selected `N`-cell branch, the source key plus `N` effective aliases can require
`N + 1` condition-cache entries even though the result contains only `N`
cells. A stack limit that admits the planner and output but is below that alias
count can produce a `ResourceLimit` solely because of duplicate cache keys.

### VP15-2: shape-mask outer capacity is released from the budget while retained

**Blocking for the shared storage-budget guarantee.**

At the start and end of `shape_hint_demand` (around lines 2848–2851 and
2961–2964), the code clears `shape_masks` and sets
`shape_mask_reservation = None`. Clearing a `Vec` retains its allocation, and
dropping the reservation releases its charged memory. A later
`install_shape_mask` call reaches `ensure_capacity` (around lines 3155–3170),
which returns without reserving when the required length fits the retained
capacity. The outer `ShapeMask` slot allocation is then physically retained
but uncharged against the shared storage budget.

The reservations owned by each mask's `indexes` vector are dropped with the
mask and do not repair this outer-vector leak. The outer vector must either
retain its reservation for as long as its capacity is retained or be replaced
with a fresh allocation before the reservation is dropped.

### VP15-3: infix reference operators lose known shape inside lazy planning

**Blocking if the array/reference profile includes nested `:`, `!`, or `~`
expressions; otherwise this needs an explicit typed refusal.**

`combine_planned_children` returns `None` for all three infix reference
operators (around lines 3057–3065), even when both operands have bounded local
reference geometry. A nested `Range` condition therefore falls back to the
current demand shape in `ShapeFrame::LazyCondition`, rather than contributing
the rectangular range shape. For example, with `A1 = TRUE` and `B1 = FALSE`,

```text
=IF({TRUE()};IF(([.A1]:[.B1]);1;0);0)
```

is planned as `1 x 1` and reads only `A1`; the reference-valued range can
provide a `1 x 2` condition and should drive two projected output positions.
The same under-admission affects a selected branch whose value is a bounded
infix range. `Intersection` and `Union` may legitimately remain unknown when
they produce a list or multiple areas, but a bounded single-area `Range` can
be resolved from metadata. If this revision intentionally excludes infix
operator geometry, it should refuse that nested case explicitly rather than
silently substitute the demand shape.

## Cleared properties

The revision does have sound structural properties within the tested direct
reference and inline-array cases:

* `ShapeFrame::LazyCondition`, `run_from`, and the scalar probe use explicit
  frame loops. The planner does not recurse through the AST or materialize a
  second evaluator.
* Lazy planning evaluates the condition before scheduling branches and installs
  masks only for selected positions. Scalar matrix emission projects one array
  cell at a time; it does not materialize an unselected branch.
* Shape growth is monotone and iteration-count bounded by expression node count
  and the configured stack limit. Shape arithmetic and mask expansion use
  checked operations, and planner cell visits charge work/cancellation through
  the shared scalar evaluator.
* Resolver reads are admitted and cancellation-checked in
  `read_reference_cell`; the cache and mask findings above are the remaining
  accounting gaps found in this bounded planner review.

**Disposition:** do not mark this planner review clear for unrestricted nested
array/reference acceptance until VP15-1 and VP15-2 are fixed. VP15-3 must be
fixed or explicitly bounded by the accepted reference-operator scope.
