# Public value and worksheet resolver API review

Reviewed at `2026-09-13T18:00:15Z` by source inspection only. No Cargo build or
runtime test was run for this review. The review covers the stable public
surface of `evaluation::value`, the public worksheet resolver, and the parent
scalar context/error types that the value API reuses. Private value-VM frame,
shape-planner, and operator behavior are outside this review; the separate
reference and VM reports remain authoritative for those paths.

The public shape is compatible with the accepted ADRs in several important
ways: the resolver is a caller-supplied synchronous trait rather than a boxed
runtime or a source/package generic; contexts carry an explicit position,
mode, and execution policy; result values are borrowed views; dimensions use
checked half-open metadata; and the worksheet adapter borrows the immutable
sheet graph while retaining one budgeted physical-run index. The worksheet
adapter's type mapping and index ownership are recorded in the
[`worksheet-resolver-review`](worksheet-resolver-review.md) report.

Two contract items require resolution before this surface is advertised as a
finished API. The reference-cell limit currently has a useful but undocumented
second meaning, and an unsupported worksheet cell is reported with the wrong
public capability diagnostic. The remaining findings are API-contract or
ergonomic follow-ups rather than claims about the private evaluator's current
formula semantics.

## Reviewed source

| file | SHA-256 |
| --- | --- |
| [`evaluation/value.rs`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value.rs) | `8220dbd9d2ac51589770eb061db03d43136cc545f5e3f810fe59aa0aa5fbaa61` |
| [`evaluation/value/references.rs`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value/references.rs) | `1f123b48b9c38b8878c0f26e58c841e08c1bbb019be5086eeda87f0bc5319fb3` |
| [`worksheet/formula.rs`](../../../../../crates/litchi-ods/src/worksheet/formula.rs) | `0b6bcb4e80875951a4a875d8f63f4964644ca07a446469bd1ce4f2aa1fc0904b` |
| [`worksheet/formula/index.rs`](../../../../../crates/litchi-ods/src/worksheet/formula/index.rs) | `4f73f3ae73e3c0e192460ef7ea209108fee0f24b0c5f27fea329f49b62741988` |
| [`evaluation.rs`](../../../../../crates/litchi-ods/src/codec/formula/evaluation.rs) | `b9bc38f50827b9c5de6ec8604d1825363c9fa45eefaf87289ac87e923bcc0f3d` |
| [`ods_formula_array_reference_evaluation.rs`](../../../../../crates/litchi-ods/tests/ods_formula_array_reference_evaluation.rs) | `6a80f05d3aa878da661d28ead10363f3859f342866f9b73dcb2241faa48d8a36` |
| [`ods_worksheet_formula_resolver.rs`](../../../../../crates/litchi-ods/tests/ods_worksheet_formula_resolver.rs) | `01e330494b779f110c6c30a1b29337e279619e899dda1cddeb41f872369eed94` |

The value evaluator was still an untracked work-in-progress file when these
hashes were captured. The hashes bind this report to the inspected snapshot;
private VM edits after that snapshot require a separate review.

## Findings requiring resolution

### VAPI-1: `max_reference_cells` also admits logical reference geometry

**Severity: contract blocker unless the public documentation is narrowed.**

[`Limits::with_max_reference_cells`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value.rs:405)
is documented as the maximum cells “fetched from one or more references.” In
the implementation, `check_reference_cells` is also called while constructing
direct references and ranges, while planning `:`, `!`, and `~`, and before
materializing a reference array. A matrix evaluation of a bare
`[.A1:.A100]` can therefore fail with `max_reference_cells = 1` before any
`Resolver::read_cell` call, and scalar projection of that vector can be refused
for its 100-cell logical extent even though it would select at most one cell.
This is a sensible admission guard against retaining or materializing a large
reference, but it is a different quantity from provider cells actually read.

The separate `reference_cells_read` counter in
[`read_reference_cell`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value.rs:4159)
does enforce the cumulative provider-read bound before each read. That guard
must remain: splitting the geometry admission from the actual-read accounting
must not reopen the multi-reference or duplicate-list loophole recorded in the
reference-operator review.

Choose and document one of these explicit contracts before release:

* rename or document this setter as a maximum logical reference extent admitted
  or fetched, and state that it can reject an unconsumed reference; or
* add a separate extent/cardinality limit and reserve
  `max_reference_cells` for cumulative resolver reads.

Add one matrix bare-range case and one scalar vector-projection case to the
limit matrix so this distinction stays reviewable. Existing tests cover the
zero direct-read limit and two eight-cell ranges, but do not state this
geometry-versus-read distinction in the public API.

### VAPI-2: `CellRead::Unsupported` is exposed as a reference-capability error

**Severity: typed-diagnostic blocker for a published resolver API.**

[`CellRead::Unsupported`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value.rs:298)
means a valid cell type outside the scalar profile. The worksheet adapter uses
it deliberately for formula-bearing cells whose cached values are inert, and
for Date, Time, and Unknown values
([`worksheet/formula.rs`](../../../../../crates/litchi-ods/src/worksheet/formula.rs:221)).
The value path maps every such result in `read_to_element` to
`UnsupportedKind::Reference`
([`value.rs`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value.rs:4462)).
That public kind displays “does not resolve references,” which is false when a
resolver was supplied and the refusal is instead about a cell value type.

Add a dedicated `UnsupportedKind::CellValue` (or a reason-bearing
`CellRead::Unsupported`) and preserve the existing `Reference` kind for a
missing reference capability or an external Source. A focused regression
should exercise a worksheet formula cache and a Date/Time/Unknown cell and
assert the typed diagnostic. This keeps operational resolver failures distinct
from formula-level errors and prevents callers from incorrectly retrying with a
different reference provider.

### VAPI-3: the evaluator context and worksheet resolver context can diverge

**Severity: resource/cancellation contract decision required.**

`worksheet::formula::Resolver::new` takes an `ExecutionContext` and retains a
clone for index construction and every later lookup
([`formula.rs`](../../../../../crates/litchi-ods/src/worksheet/formula.rs:149)).
`value::Context` separately takes the execution context used by
`value::evaluate` ([`value.rs`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value.rs:204)),
but `evaluate` has no association check between the two. The documentation
example passes the same context, yet a caller can construct the resolver with
one budget/token and invoke evaluation with another. In that case index and
lookup work, index memory, and lookup cancellation are charged through the
resolver's retained policy, while VM storage/work and the final cancellation
fence use the evaluation context. The result is safe against the evaluator's
own final fence, but the operation's caller-visible budget and cancellation
lineage are split.

Either make the same context lineage an explicit documented precondition with a
regression for mismatched contexts, or provide an API association that can
validate it. Do not infer identity from equal numeric limits: unrelated
contexts can have the same limits but different budgets or cancellation
tokens. This is a boundary decision for the public API; the worksheet index's
retained context and its per-lookup checks are otherwise sound.

### VAPI-4: `source_version = None` needs a stronger coherence contract

**Severity: P2 contract ambiguity.**

The trait's default `source_version` returns `None`
([`value.rs`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value.rs:342)),
while `evaluate` documents that all reads are against one immutable source
version and only performs before/after fencing when an identity is supplied
([`value.rs`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value.rs:1176)).
Allowing stateless synthetic providers without a positional source is
reasonable, but the public contract should say that `None` is valid only when
the provider guarantees one coherent immutable set of sheet names, extents,
and cell values for the result lifetime. It should also state that
`sheet_index`, `sheet_name_at`, and `sheet_count` are mutually consistent over
one ordered set. Otherwise a mutable provider can opt out of the only runtime
source-change check while still returning borrowed data.

The worksheet adapter satisfies this requirement by borrowing immutable sheets;
this finding concerns the generic trait contract, not that adapter's source
model. A versioned custom provider remains the preferred path when its source
can change.

### VAPI-5: borrowed view equality is pointer identity without documentation

**Severity: P2 ergonomics/semantic contract.**

`Value` derives `PartialEq`, but `ArrayView` compares its cell slice with
`std::ptr::eq` and `ReferenceListView` compares its records slice the same way
([`value.rs`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value.rs:581),
[`value.rs`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value.rs:664)).
`ReferenceView`, by contrast, has structural derived equality. Two independent
evaluations containing identical array or reference-list data therefore compare
unequal solely because their retained vectors are different allocations, while
two views into one allocation compare equal. The public docs describe borrowed
views and source order but do not state identity-based equality.

Either implement structural equality over the bounded cells/records or remove
the equality implementation from the view-bearing public values. If pointer
identity is intentional for an O(1) operation, document it prominently and
provide a named semantic comparison path so `Value` is not mistaken for a
normal data-bearing value. A regression should compare two separately
evaluated equal arrays and reference lists.

### VAPI-6: `Evaluated` has no explicit owned conversion

**Severity: P2 ergonomic follow-up, not a lifetime-safety defect.**

`Evaluated` intentionally retains the expression, resolver, and its result
allocations through a shared borrow lifetime and exposes only `Value`,
`ArrayView`, and reference views
([`value.rs`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value.rs:713)).
There is no `to_owned` or `into_owned` operation for an array or reference
result. A caller that needs to retain a matrix after dropping the worksheet
resolver must manually walk and copy every cell; a caller cannot reconstruct a
fully owned reference view from the public API without separately preserving
the parsed `Reference` and area metadata.

ADR 0003 requires an explicit conversion when a borrowed zero-copy input is
made persistent. Add a budgeted owned conversion when the owned representation
is defined, or narrow the API documentation to state that this is an
ephemeral-inspection-only result and that callers must perform their own
copying. This is separate from the current correct borrow checker boundary.

## Lower-priority API notes

* `EvaluationOptions` has a private `text_case` field and the only
  `TextCase` variant is `Sensitive`. `Context::with_options` is therefore
  callable but cannot currently express a setting different from
  `Default`. Either add a builder when another policy is supported or narrow
  “comparison/conversion options” to the one currently implemented policy.
* `SheetExtent::new` accepts a zero row or column count, while the concrete
  worksheet resolver rejects either as `InvalidExtent` and `Shape` rejects
  zero dimensions. The generic `Resolver` contract should state whether an
  empty extent is valid and how it is represented; this is currently a
  consistency/documentation issue rather than an observed unsafe path.
* `Area::from_bounds` intentionally validates only nonempty half-open numeric
  bounds. It is a value descriptor, not a resolver lookup, so callers should
  not read its arbitrary sheet indices as proof that a provider contains those
  sheets. The evaluator and worksheet adapter perform the actual extent/name
  checks.

## Confirmed boundaries

The public surface does satisfy the following reviewed constraints:

* `Resolver` uses static dispatch and `&self` methods; it exposes no archive
  handle, package ID, raw physical ID, ambient I/O, or boxed trait object.
  `Position`, `Shape`, and `Area` keep construction checked at the semantic
  boundaries, and `Value`/`CellRead` are non-exhaustive data-bearing enums.
* `Context` keeps the caller's execution reference and explicit mode/position.
  `Limits` provides finite local maxima, while `EvaluationFailure` preserves
  typed cancellation, resource, allocation, execution, and source-version
  outcomes. The top-level value path has entry, source-version, and final
  cancellation fences in the inspected snapshot; this report does not
  re-review the private lazy/matrix implementation behind those fences.
* The worksheet resolver borrows `&[Sheet]` or `Snapshot::sheets()`, keeps
  repeated rows/cells as physical runs, retains one exact index reservation,
  checks the retained context during lookup, and returns borrowed text. It
  refuses formula caches and Date/Time/Unknown values instead of silently
  treating them as calculated scalars. Its explicit finite extent is separate
  from stored worksheet repetition bounds.
* External `Reference::Source` values remain an explicit unsupported
  capability; the resolver contract does not perform network or source refresh
  work. This is consistent with [ADR 0001](../../../../adr/0001-priorities-and-api-layers.md),
  [ADR 0003](../../../../adr/0003-snapshots-edits-and-patches.md),
  [ADR 0004](../../../../adr/0004-semantic-api-design.md),
  [ADR 0005](../../../../adr/0005-io-memory-and-performance.md), and the
  array/reference requirements in [`spec-review.md`](../spec-review.md).

## Acceptance regressions

Before publication, retain focused tests for the two contract blockers and
the documented follow-ups:

1. Evaluate an unconsumed large reference and a scalar projection under a
   small `max_reference_cells`; assert whether geometry admission or actual
   provider reads are being limited, while retaining the cumulative-read
   test for two ranges and duplicate list entries.
2. Return `CellRead::Unsupported` for a formula cache and Date/Time/Unknown
   through the worksheet adapter; assert a cell-value capability diagnostic,
   not `UnsupportedKind::Reference`.
3. Exercise resolver/evaluator contexts with distinct budgets and cancellation
   tokens, and either reject the mismatch or document the split policy.
4. Compare separately allocated equal arrays and reference lists, and add a
   compile-level lifetime example for the chosen owned-conversion or
   ephemeral-result contract.

Passing existing integration tests is useful evidence for the physical index,
source fencing, and cumulative read counter, but does not resolve these public
contract decisions or inspect the private VM paths excluded above.
