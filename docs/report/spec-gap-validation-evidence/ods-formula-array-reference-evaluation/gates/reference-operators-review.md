# Non-lazy reference operator and projection review

Reviewed at `2026-09-13T16:48:10Z` by source inspection only. No Cargo build or
runtime test was run for this review. The review is against the accepted
reference rules in [`spec-review.md`](../spec-review.md), especially the
`Reference`/`ReferenceList` distinction, ordered sheet traversal, explicit
scalar projection, and aggregate resolver-cell limits.

## Reviewed source

| File | SHA-256 |
| --- | --- |
| [`evaluation/value.rs`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value.rs) | `da91667d01a7d7f2512b970d2ade5aaa2f7fdaa0ec660d436345c62cf758a779` |
| [`evaluation/value/geometry.rs`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value/geometry.rs) | `ad94bd57f5aee96e90b1d9ec79217b4689c5bfbe422c3a4d62e278da7b22ff7d` |

The first file was an untracked work-in-progress snapshot when fingerprinted;
the hash binds the observations below to that snapshot.

## Disposition

The bounded physical-sheet traversal and the ordinary single-reference scalar
projection are directionally correct. The reference substrate is still blocked
for acceptance by list identity, range geometry, reference-cell accounting,
and explicit projection behavior described below. Green tests that exercise one
surviving intersection or one 3-D range do not cover these cases.

The findings in the next section were recorded against the monolithic
`value.rs` snapshot shown above. The extracted reference-operator and
worksheet-resolver follow-up at the end of this report supersedes those
statuses where it records a fix in the current source. The historical text is
retained so the original review remains auditable.

## Initial snapshot findings (historical)

### `!` loses list-entry identity and retains empty records

`intersect_ranges` marks a result as a list when either operand is a list, but
starts one derived `RuntimeReference` and calls `push` for every surviving
left/right pair. `push` appends areas to the last record; it never creates one
record per surviving pair. For example,

```text
([.A1]~[.B1]) ! ([.A1]~[.B1])
```

has two surviving combinations (`A1` and `B1`) and must expose two ordered
list entries. The current result exposes one list record containing two areas.
When every combination is empty, the seeded empty record remains. Matrix mode
therefore exposes a reference/list containing an empty entry, while scalar
projection turns the empty area set into `#NULL!`; the accepted profile requires
empty combinations to be omitted and an all-empty intersection to be a formula
error. This also affects `AND`/`OR`, which can aggregate the phantom empty set.

Required regressions:

* assert two records and their order for the two-list example above;
* assert an all-empty intersection is an error in both matrix and consuming
  contexts, with no empty public record.

### `:` does not produce the smallest cuboid for list operands

`combine_range` computes a `Rect::bounding` and sheet span independently for
each raw left/right area pair, then appends all pair results into one derived
reference. Under the accepted profile, `:` over a `ReferenceList` must extend
one range over every element and return the smallest inclusive 3-D cuboid. For
example,

```text
([.A1]~[.C3]) : [.B2]
```

should produce one `A1:C3` cuboid. The current implementation produces two
overlapping rectangles (`A1:B2` and `B2:C3`) in one `ReferenceView`. Its
`planned_cells` check also sums pairwise rectangles, so overlapping or duplicate
list entries can be refused even when the final cuboid is within the limit.
The range implementation needs one global checked bound (including the global
sheet span) before emitting one ordered cuboid, with a regression using two or
more list entries and a 3-D list case.

### `max_reference_cells` is not an aggregate fetched-cell limit

The public limit is documented as the maximum cells fetched from one or more
references, but the current accounting has three holes:

* `reference_areas` admits an `Address::Cell` without checking
  `max_reference_cells`; scalar `=[.A1]` can still call `read_cell` when the
  limit is zero.
* `apply_sequence` (`AND`/`OR`) scans every area and calls `charge_cell_work`,
  which charges scalar work but no cumulative reference-cell counter. Thus
  `AND([.A1:.A8];[.B1:.B8])` can read 16 cells with
  `with_max_reference_cells(8)` because each syntactic reference passes its
  own eight-cell admission.
* `append_set` and `intersect_ranges` retain multiple areas without a checked
  aggregate cell count. Later sequence consumption can read all of them.

`materialize_areas` checks one rectangular area, and ordinary 3-D address
construction checks its one cuboid before emitting per-sheet areas; those checks
do not close the multi-reference paths. Every actual provider read, including a
direct projection and duplicate list entries, needs a checked cumulative
admission policy before the read. Add regressions for a zero-cell direct limit,
two references whose combined count exceeds the limit, and a duplicate/list
case so the chosen duplicate-counting policy is explicit.

### Public explicit projection cannot dereference a reference

`Evaluated::implicit_intersection` is documented as returning the scalar result
of explicit implicit intersection, but `RuntimeValue::implicit_intersection`
only computes an area candidate. For `RuntimeValue::Areas`, a selected area
returns `Value::Error(ScalarError::Reference)` and no `Resolver::read_cell` is
performed. The evaluator's private `project_area_value` does read a cell, so
scalar evaluation of the original expression can appear correct while the
public matrix-result projection API is unusable for every reference.

Either retain the resolver capability in the evaluated projection object or
make projection an explicit resolver-taking operation. In either design,
`B2` with `[Data.A1:.A3]` must read `Data.A2`, a 2-D rectangle must retain the
documented ambiguity error, and a 3-D cuboid must not silently select its first
sheet. Add a public API regression; the current scalar-vector tests do not call
this method.

### Final cancellation fence is missing around provider calls

The top-level `evaluate` checks cancellation before `source_version`, and the
large loops periodically check through `charge_cell_work`, but it does not
check after `finish_root` or after the final `source_version` call before
publishing `Evaluated`. In particular, `project_area_value` charges scalar work
and then calls `Resolver::read_cell` without a final local execution check. A
caller-owned resolver can cancel the shared context during that read and return
success; the missing final fence then lets a successful result escape.

Add a check after the last potentially blocking resolver operation and before
publishing the borrowed result. A resolver that cancels during the final cell
read should produce typed `Cancelled`, including for a one-cell scalar
projection. The existing pre-cancelled test only proves the entry fence.

## Finite extents and cleared geometry paths

Whole-row and whole-column endpoint construction obtains `SheetExtent`, checks
the endpoint against it, computes checked cell counts, and rejects a 3-D span
whose end is outside `sheet_count`. Intermediate 3-D names are obtained in
physical resolver order; reversing endpoint order is normalized to that order.
The worksheet adapter additionally rejects out-of-extent cell reads, so its
current concrete path is bounded and returns `#REF!` for such coordinates.

The generic evaluator does not consult `sheet_extent` for an explicit cell
endpoint: `endpoint_rect` creates `[.Z999999]` from syntax alone and leaves the
decision to `read_cell`. A custom resolver that returns `Empty` for an
out-of-extent coordinate would therefore violate the accepted finite-grid
profile. Document `read_cell` as responsible for this invariant or preflight
explicit coordinates in the evaluator; retain an adapter regression proving
the refusal.

For scalar projection, the current private path is consistent with the chosen
profile: one-sheet row/column vectors use the caller's varying coordinate even
when the target sheet differs; a multi-sheet cuboid selects only the caller's
physical sheet plane; a 2-D rectangle with multiple candidates returns
`#N/A`; and `ReferenceList` is refused instead of selecting its first entry.
Matrix projection refuses multi-area/3-D references rather than flattening
sheets. These properties do not repair the public projection method or the
operator findings above.

Source-version fencing is present for successful top-level evaluations: the
before/after identities are compared and availability changes are refused.
Provider errors return before the after sample, which should remain an explicit
error-precedence decision. This review found no separate source-version bypass
for an immutable resolver; the missing cancellation fence is independent.

## Suggested focused regression set

1. Two-list `!` with two surviving pairs; all-empty `!` in matrix and scalar
   consumers.
2. List `:` over separated cells and over endpoints on different sheets;
   assert one global cuboid and checked final-cell count.
3. `max_reference_cells(0)` for a direct scalar cell; two eight-cell ranges
   passed to `AND` under a limit of eight; and an over-limit duplicate list.
4. Matrix evaluation followed by public `Evaluated::implicit_intersection` for
   a cross-sheet column vector, an ambiguous 2-D rectangle, and an interior
   3-D plane.
5. A resolver that cancels during the final `read_cell`, plus explicit cells
   outside its finite `SheetExtent`.

For the historical snapshot, until these cases were covered and the four
semantic/resource issues were corrected, its passing test subset was only
partial validation rather than acceptance of the reference substrate. The
current disposition is in the follow-up below.

## Follow-up: extracted reference operators and worksheet resolver

Reviewed at `2026-09-13T17:34:29Z` by source inspection only. No Cargo build or
runtime test was run for this follow-up. The source snapshot was fingerprinted
as follows:

| File | SHA-256 |
| --- | --- |
| [`evaluation/value/references.rs`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value/references.rs) | `1f123b48b9c38b8878c0f26e58c841e08c1bbb019be5086eeda87f0bc5319fb3` |
| [`evaluation/value/geometry.rs`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value/geometry.rs) | `ad94bd57f5aee96e90b1d9ec79217b4689c5bfbe422c3a4d62e278da7b22ff7d` |
| [`worksheet/formula.rs`](../../../../../crates/litchi-ods/src/worksheet/formula.rs) | `0b6bcb4e80875951a4a875d8f63f4964644ca07a446469bd1ce4f2aa1fc0904b` |
| [`worksheet/formula/index.rs`](../../../../../crates/litchi-ods/src/worksheet/formula/index.rs) | `4f73f3ae73e3c0e192460ef7ea209108fee0f24b0c5f27fea329f49b62741988` |
| [`evaluation/value.rs`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value.rs) (context only) | `f0f4c792fbc7478ebd9163d3e73bceacd4a205d69e29703f8310dd309960c53e` |

The `value.rs` hash is included only to identify the VM snapshot in which the
read-accounting and cancellation helpers were observed; the VM was changing
under the implementation owner and was not re-reviewed wholesale here.

### Current operator disposition

* `:` now computes one checked global rectangle and physical sheet span across
  all operand areas. It admits the complete cell count before retaining output,
  checks every sheet in the span against its finite extent, and emits one
  non-list reference with one physical area per sheet. The former pairwise
  range-geometry finding is closed for this snapshot.
* Cell, row, column, and multi-sheet endpoint construction now performs finite
  extent admission. The multi-sheet path validates every intermediate physical
  sheet before retaining areas. The former deferred-out-of-extent finding is
  closed for this snapshot.
* `!` plans and emits by left-record/right-record pair, so one surviving pair
  retains all of its intersecting physical planes as one list record. Empty
  pairs are skipped; when no pair survives, `ScalarError::Null` is returned.
  That is a formula Error and therefore satisfies the accepted all-empty rule;
  the old phantom-empty-record finding is closed.
* The former list-identity blocker is closed in this snapshot. `!` now derives
  each record's contiguous physical-plane slice from the sum of its public
  sheet extents, checks that every slice stays within the shared physical-area
  vector, and verifies complete coverage after both planning and emission.
  This preserves duplicate records and prevents a multi-sheet public cuboid
  from claiming a nested plane belonging to another record. The revision's
  focused duplicate and nested-3-D tests are reported as passing by the owning
  runtime gate; this review did not run them.
* The `!` planning and emission scans are bounded by the retained
  record/area vectors. Each record-extent summation is preceded by a work
  charge sized to its public-area count, and every physical left/right
  comparison, including unsuccessful sheet or rectangle matches, is charged.
  `combine_range` now obtains each output sheet name directly from the
  resolver's ordered `sheet_name_at(index)` contract, with one charged unit per
  physical sheet and no operand-area rescan. I found no remaining
  undercharged scan in the extracted operator module; the two-pass `!` work is
  explicit and bounded.
* The current VM has a cumulative `reference_cells_read` counter in
  `read_reference_cell`, with admission before each provider read and
  cancellation checks before and after that read. The operator preflight also
  checks retained reference cell/area totals. This closes the historical
  direct/aggregate `max_reference_cells` finding for the observed snapshot;
  the VM was not otherwise re-reviewed here.
* The current top-level value path checks the retained execution context after
  the final source-version probe. This closes the historical final-cancellation
  fence finding for the observed snapshot. Publication still depends on the
  broader VM gates owned by the implementation review.
* The former public `Evaluated::implicit_intersection` method is absent from
  the current API. The old method-specific finding is therefore superseded;
  this review does not claim that a replacement public projection operation
  exists.

### Worksheet adapter and physical-run index

The current adapter borrows the caller's immutable `&[Sheet]` (and
`Snapshot::sheets()`), stores borrowed names plus source indices, and returns
borrowed cell/text data. Its index retains the caller's `ExecutionContext` and
one exact metadata reservation. Sizing, population, name sorting/comparison,
and binary searches charge work and check cancellation. Repeated rows and
cells remain physical run descriptors; lookup does not expand logical cells.
The heap comparator has a deterministic UTF-8/name-plus-source-index order,
while duplicate exact names are rejected without hiding equal names behind the
tie-break.

`read_cell` enforces the caller's finite `SheetExtent` before returning a cell.
Absent physical cells are `Empty`; formula-bearing cells refuse their cached
values; finite Number/Currency/Percentage and Boolean map to numeric/logical
values; Text remains borrowed; Date, Time, Unknown, and non-finite numbers map
to typed refusals/errors. This keeps worksheet caches from becoming an
implicit formula evaluator and preserves source identity without cloning.
The current `sheet_count` and `sheet_name_at` methods also check the retained
execution context, closing the cancellation gap recorded in the separate
earlier adapter review.

The adapter's remaining acceptance boundary is integration: the current source
inspection does not establish that every public VM path preserves these
borrowed values and typed refusals under all matrix/reference operators. The
existing integration gates should retain focused cases for repeated runs,
formula-cache refusal, Date/Time/Unknown refusal, finite extents, duplicate
sheet names, cancellation during lookup, and Text lifetime. No additional
physical-run expansion or reservation leak was found in this inspection.

### Focused regression retained

Retain reference-list intersection tests for `([.A1]~[.A1])![.A1]` and
`([S1.A1:S3.A1]~[S2.A1])![S2.A1]`; each should assert two records, one area per
record, and preserved order. Keep the existing all-empty `#NULL!` assertion,
and retain a multi-plane duplicate/nested-cuboid case so private identity
cannot regress while ordinary 3-D grouping tests remain green. The owning
revision-8 runtime gate reports the duplicate and nested-3-D cases passing;
this review did not run that gate.
