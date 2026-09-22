# ODS financial value adapter resource review

This review covers the resolver-backed value adapter, its scalar façade, and
the financial integration in the value evaluator. It applies ADR 0005's
bounded work, storage, read, and cancellation rules and ADR 0006's typed
failure and publication rules. It is bound to the source snapshot identified
below. No production or test file was changed for this review.

The disposition is **PASS for the reviewed resource and cache boundaries**,
with three non-blocking implementation advisories recorded below. This is a
resource/cache verdict only; numerical CUMIPMT/CUMPRINC and scalar-kernel
correctness remain covered by their separate review. A previous
projected direct-rate NPV read-count diagnostic is retained as historical
evidence; the current focused limits run passes after the projection fix.

## Reviewed inputs

| Input | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation/value/financial.rs` | `8e471abe6c842369f8df0e04079121951720b82dfbdce541d1a331284fdfd8ef` |
| `crates/litchi-ods/src/codec/formula/evaluation/financial.rs` | `93498020996fa9c159a026466553e8f464d9b5a32272723acd2060e69ec6fa0d` |
| `crates/litchi-ods/src/codec/formula/evaluation/financial/kernel.rs` | `f7147ef31ba1642cf9f5e14c4006d7039fcc1ce7c1977d818b23c801c0807af4` |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `110d63d6a80e6294861c7c423f597a7cb46d23f6300d7ccb95c81c97aa31cba4` |
| `ods-formula-financial/contract.md` | `fbe2283886326582b0e34fd277b38e6fb3e1bdeb2ecfbec5522da0f6f1467e4b` |
| `ods-formula-financial/ods_formula_financial_limits.rs` | `b210d00653e276a7ba62ddc3fccc78a9a69552d1ece382259229ced37282868f` |

The source was reviewed in the isolated financial development checkout at
checkpoint `ab93dec951`; the hash table is the authoritative binding because
the financial source is staged there while the parent checkout carries the
owned evidence files.

## Resource boundaries that pass review

The value adapter performs pseudotype and known-shape checks before scanning a
financial argument. `NumberSequence` and `DateSequence` reject a known
reference list, and XNPV's `Reference | Array` arguments reject a known list.
These paths produce a formula value with zero resolver reads. Computed
children are allowed to evaluate in order to establish their runtime type,
which is distinct from catching a provider failure and turning it into a
shape result.

Complete sequence arguments are scheduled with complete descriptors through a
projected matrix or lazy `IF`. Scalar rates, guesses, controls, and periods
retain projected position semantics. The financial cache classifier admits
stable literals and direct references according to that distinction and
conservatively declines computed sequence branches. The criterion/full-
argument path propagates complete financial sequences; position-sensitive
scalar children such as MUNIT remain outside that propagation. A cache entry
contains only a completed scalar result or formula-error payload, never a
partial reducer or typed evaluator failure.

Reference areas are scanned in area, sheet, row, and column order. Each
physical cell charges work before `read_reference_cell`; the read path checks
the reference-cell limit and cancellation before and after the resolver. Text
is inspected through `read_to_element` and stays borrowed while byte work or
numeric conversion is charged. The scans continue after retained formula or
generated errors, so a later provider, source, cancellation, or resource
failure remains typed and supersedes the retained formula value.

NPV and FVSCHEDULE use fixed reducer state and stream their complete sequence
inputs. XNPV scans Values first into bounded numeric slots, then scans Dates
and consumes paired slots without rereading Values. IRR uses bounded numeric
scratch; XIRR uses bounded value/date marker slots followed by bounded numeric
scratch; MIRR uses fixed reducer state. Array, area, slot, and output buffers
use checked limits and fallible reservations. Numeric root calls in the value
adapter hold the solver working-state reservation through the solver result.
The reducer bridge preserves exponent-span failures as typed memory failures
and maps ordinary numerical failures to formula errors. Cumulative functions
use fixed scalar state and charge their checked period loops.

The focused evidence reported with this snapshot includes the financial
streaming, shape, typed-precedence, XNPV invalid-rate, cancellation, and
projected-cache cases; the current limits target is reported as 23/23 passing.
The parent owns the full integration and lint receipts. This review did not
run Cargo.

## Advisories and historical diagnostics

### Scalar IRR/XIRR has no current unreserved solver allocation

The scalar façade's `apply_irr` and `apply_xirr` call the bounded root kernels
without acquiring `financial::solver::WORKING_STATE_BYTES`, while
`apply_rate` and the resolver-backed value adapter do reserve it. The scalar
façade passes fixed one-element arrays. `IRR`'s cash-flow validation charges
and checks the sign classes before `solve_problem`; one value cannot contain
both a positive and a negative cash flow. `XIRR` first checks equal/non-empty
length, then performs the same sign validation (and date validation) before
`solve_problem`. Both scalar paths therefore return a formula number error
before the root kernel can construct its two `LogSum` accumulators. There is
no current input-sized unreserved allocation or reservation blocker under the
scalar contract. If scalar sequence admission is broadened later, reserve
the solver state at that new call boundary and retain the reservation through
the result.

### XIRR retains marker and numeric buffers at the same time

`crates/litchi-ods/src/codec/formula/evaluation/value/financial.rs::apply_xirr`
allocates and fills `value_slots` and `date_slots`, then allocates
`values_scratch` and `dates_scratch`, and only after the pairing loop drops the
marker buffers. The marker slots are necessary metadata: they preserve the
original positions needed for equal-length checks, one-sided skip detection,
and source-order formula/generated error precedence. The scalar
`first_admitted`/sign flags are constant-size metadata. The compact numeric
pairs are necessary solver input, but `scratch_for` reserves capacity from the
original value/date geometry rather than the eventual admitted-pair count.
All four vectors have checked reservations, so the state is bounded and typed
storage failure is handled. The ordering nevertheless raises the peak to two
full-capacity marker vectors plus two numeric vectors, in addition to the
solver state. The nearby comment is true at the solver boundary, after the
`drop` calls, but not during numeric-pair allocation. A future memory-profile
fix should use a representation that does not retain four full-capacity
vectors simultaneously, while preserving original pairing/error order and
avoiding resolver rereads. This is a bounded but avoidable storage spike that
can reject an input under a budget that would fit the intended representation;
it does not violate the current checked-accounting boundary.

### Projected direct-rate NPV reread diagnostic (fixed)

The focused projected NPV case with two distinct scalar rates and one complete
three-cell Values reference is intended to read two rate cells plus one
complete Values scan per result, for eight reads. An earlier run reported
sixteen reads for the direct-rate branch. The adapter projection fix is now in
the reviewed `value/financial.rs` snapshot: the focused limits target reports
23/23, and the direct-rate case records two rate reads plus one complete
Values scan per result. The invariant-rate cache and position-sensitive MUNIT
cases remain covered. This historical failure should stay as a regression
test, but it is no longer an acceptance blocker.

### Minor AST accounting advisory

`matrix_argument` calls `is_array_node`, which walks an arbitrary chain of
parenthesized AST nodes without charging each step. The surrounding function
argument scheduling charges the child count, and all runtime cell and value
loops are charged, so this is not an observed unbounded data allocation. Add
a bounded AST work charge if parenthesis depth is not already covered by the
general expression budget.

## Required handoff conditions

Retain focused evidence for: zero-read known shape refusals; one-pass area
order and borrowed text; formula Error followed by a late typed provider,
cancellation, source, or resource failure; Values-before-Dates XNPV traversal;
bounded IRR/XIRR scratch and solver work; reservation drop on every error
path; projected complete-sequence cache behavior; and position-sensitive
scalar parameters. Refresh this report after any source change affecting
these boundaries. Numerical CUMIPMT/CUMPRINC and scalar kernel stability
changes are outside this adapter resource verdict and require their own
kernel review.
