# ODS financial contract resource review

This is a design review of the twenty-function financial contract in
[`contract.md`](contract.md). It addresses the bounded evaluator shape required
by ADR 0005 and the typed-error, provider, and publication rules in ADR 0006.
No production, test, or Cargo files were changed. The contract is a useful
resource boundary, but it is not ready for a support claim: several numerical
and pseudotype choices deliberately remain open and must be fixed before an
implementation can be reviewed as conforming.

## Reviewed inputs

| Input | SHA-256 |
| --- | --- |
| `ods-formula-financial/contract.md` | `099581eff819d57a4e2e23dfdc4d04846bc49597d3654bd4d40b572582ecf1b1` |
| `docs/adr/0005-io-memory-and-performance.md` | `34a6148a8fe77b3e90212810996667b654e25209fb49c56aaecd8d8fbe83f770` |
| `docs/adr/0006-validation-security-and-compatibility.md` | `b686465f342b2e051f856f094e38c2333aa05a2768bb255fc5b223cbba1f3381` |

The current contract also records the exact 54-entry §6.12 inventory and the
20 entries in this slice. It correctly adopts finite hierarchical work, storage, reference,
allocation, and cancellation budgets; fallible reservation; source-version
fences; explicit providers; and no ambient clock, I/O, or external rate
service. The resource boundary is implementable if the requirements below are
made normative and the listed semantic choices are resolved.

## Streaming reducer boundaries

The implementation should classify each argument before evaluation and keep
the sequence-bearing arguments complete. The following state shapes satisfy
the contract without retaining a range of cell objects or text:

| Operation | Required evaluation shape |
| --- | --- |
| `FVSCHEDULE` | One pass over the ordered `NumberSequence`; retain only the running product and a fixed formula-error record. |
| `NPV` | One pass over every `NumberSequenceList` argument, each area in occurrence order and each area row-major; retain the running index and sum. Do not build a combined values vector merely to apply the discount exponent. |
| `XNPV` | One pass over paired admitted values and dates after the contract fixes pair traversal and error order. If source-order evaluation requires separate argument scans, use bounded numeric scratch as described below rather than rereading a provider implicitly. |
| `MIRR` | A one-pass positive/negative accumulator is possible after `Array` admission and element conversion are pinned. Text and Empty remain ignored as required; no cell or text vector is needed for the equation. |
| `IRR` | A bounded numeric cash-flow vector is required for repeated root evaluations unless the contract explicitly permits replaying the complete source under a stable snapshot and a charged read budget. Replay must not be an accidental consequence of a reducer callback. |
| `XIRR` | A bounded vector of admitted `(value, date)` numeric terms is required for repeated date-weighted root evaluations, with the same source-order and error policy as `IRR`. |
| `CUMIPMT` and `CUMPRINC` | Iterate the checked `Start..End` period domain with fixed numeric state. Charge each period and check cancellation; do not allocate one record per period. |
| `RATE` | Fixed-size scalar root state; its scalar arguments remain projected and position-sensitive. It still needs the same finite iteration and cancellation profile as `IRR`/`XIRR`. |

The remaining annuity functions are scalar numerical kernels. They do not
materialize a sequence, but their computed scalar arguments may themselves
perform bounded evaluation and must retain the ordinary source, work, and
cancellation fences.

`NPV` has a variadic sequence-list input. Its argument order, ReferenceList
area occurrence order, sheet order, and row-major cell order are part of the
numeric exponent and formula-error precedence. A complete descriptor may be
retained; a cell vector is not required. `FVSCHEDULE` has the same streaming
rule even when a schedule is a direct reference. `XNPV` cannot claim a
one-pass implementation until the contract defines whether values and dates
are paired lockstep or are evaluated as two source-order arguments. The
implementation must not silently choose a traversal that changes which late
provider failure wins.

All physical sequence cells are charged and checked before inspection,
including Empty or Text cells that the pseudotype later skips. A direct
resolver Text value stays borrowed through the conversion decision. Numeric
state may retain only finite converted numbers and a fixed error record; it
must never retain borrowed cell wrappers or cloned text.

## Bounded IRR/XIRR scratch

`IRR`, `RATE`, and `XIRR` need a numerical profile in addition to the resource
profile. For `IRR`, the recommended representation is a fallibly reserved
`Vec<f64>` containing admitted finite Number values, plus fixed state for the
first formula Error, sign presence, and generated conversion/domain errors.
For `XIRR`, use a fallibly reserved vector of fixed-size value/date pairs, or
two equally bounded numeric vectors whose paired indexing is explicit. Store
no `ScalarValue`, `TextValue`, resolver handle, or source slice in either
buffer.

The sequence descriptor must be checked before reservation. A known reference
geometry supplies an exact upper bound; a computed array uses the existing
array-cell ceiling. The reservation must account for checked element size and
the old-plus-new capacity when a vector grows. Unknown or overflowing counts
return a typed storage/allocation failure before unbounded growth. Every
reservation is released after the vector is dropped, including formula-error,
nonconvergence, cancellation, and typed-provider exits.

The initial scan must continue after a formula Error so a later resolver,
source, cancellation, allocation, or budget failure can supersede it. It must
also finish the admitted paired scan needed to establish size and sign rules;
no early positive/negative-sign shortcut is allowed. If a sequence cell is
skipped by conversion, its read and work charge still applies.

Repeated root passes operate only on the bounded numeric buffer. Each
iteration and each term evaluation charges finite work with checked arithmetic
and checks cancellation at a bounded interval. A work-limit refusal is typed
and cannot be converted to `#NUM!`, even when a formula Error was retained.
The caller's Guess is an initial value only and cannot raise the hard
iteration, term, or total-work ceiling. A fixed implementation profile must
state whether an iteration charge is made before the iteration's first term or
incrementally; either way, a checked overflow must fail before the counter
wraps.

`XIRR` date conversion must reuse the completed DateSequence profile: direct
ReferenceList mismatch is rejected before reads, admitted cells retain the
prescribed date order, and date formula Errors remain values while the
complete required scan continues. Values and dates must have equal admitted
length, and the implementation must not pair two independently projected
scalar intersections under a matrix caller.

## Shape, conversion, and zero-read requirements

The evaluator should use three explicit scheduling classes:

1. Complete `NumberSequence`, `NumberSequenceList`, and `DateSequence`
   arguments are visited as complete descriptors for every scalar result.
   `NPV` consumes every complete list argument; `IRR`, `FVSCHEDULE`, and
   `XIRR` retain their complete sequence input.
2. Scalar `Number`, `Integer`, `Guess`, `Rate`, `Payment`, `PayType`, `Type`,
   and other controls retain projected position semantics under a matrix or
   lazy `IF` caller. A scalar financial function result remains scalar even
   when a complete sequence is rectangular.
3. Explicit `Array` arguments use the array profile selected by the contract;
   they are not silently changed into `NumberSequenceList` arguments.

Known wrong shapes must be refused before resolver reads. In particular,
`NumberSequence` and `DateSequence` reject a known `ReferenceList`, and `XNPV`
rejects a known `ReferenceList` because its signature is `Reference | Array`.
The refusal is a formula shape/type result with zero cell reads; it still
charges bounded descriptor geometry and must obey `max_reference_cells` and
array limits. A computed child may read while producing its runtime type, as
the existing evaluator contract permits. `MIRR(Array Values)` is not ready for
this guarantee until its reference and array admission is decided.

Integer conversion uses checked finite truncation toward zero. No path may use
an unchecked float-to-integer cast for period, type, payment-type, or
compounding controls. Scalar Text conversion scans borrowed text under its
text/work budget and does not clone it to test numeric syntax.

## Formula-error and typed-failure precedence

Formula Errors are values. Each reducer needs separate fixed state for the
first formula Error in conceptual source order and for generated formula
errors. It must continue the complete admitted scan after either ordinary
formula error so a later typed failure is observable. The established
precedence should then publish the first applicable formula Error or generated
result according to the final financial profile.

Resolver `Unsupported`, `ResourceLimit`, `Allocation`, `Cancelled`,
`SourceChanged`, and source-version failures abort immediately as typed
evaluator failures. They supersede all retained formula or generated errors,
are never cached as formula values, and cannot be caught by `IFERROR` or
`IFNA`. This includes a provider failure in a later ReferenceList area, a
later XIRR date after an earlier value Error, and a failure during a repeated
root scan.

Shape refusal is a distinct generated formula result: it is allowed to avoid
resolver reads when the wrong descriptor kind is already known. It must not be
implemented by catching a provider `Unsupported` or `ResourceLimit` from a
computed child. The evaluator needs an internal distinction between “known
wrong shape” and “provider failure while probing shape”.

The contract still needs an explicit cross-argument order for paired XIRR and
XNPV values/dates. Until that is fixed, a lockstep implementation and a
values-first/dates-second implementation can disagree when one argument holds
a formula Error and the other later produces a typed provider failure. The
same decision is needed for a scalar Guess or Rate Error following a sequence
Error. Source order must be observable and tests must cover the late-failure
case.

## Work, cancellation, source, and drop order

* Every reference cell read charges cell work before coordinate arithmetic and
  before the resolver call, checks the reference-cell limit, and checks
  cancellation before and after the provider. Successful reads are counted
  only through the existing read path. The same rules apply to values and
  dates in XIRR/XNPV and to cells later skipped as Empty or Text.
* NPV/FVSCHEDULE/XNPV sequence terms charge the arithmetic work associated
  with conversion and accumulation. CUMIPMT/CUMPRINC charge every period;
  checked period cardinality and work arithmetic prevent a huge integer range
  from becoming an unchecked loop.
* IRR/RATE/XIRR charge every root iteration and its term evaluations, check
  cancellation at least once per iteration and at a bounded term interval,
  and check before publishing the result. The profile must define the
  interval; an unbounded numerical loop is not acceptable.
* Source-version checks surround the complete evaluation and precede
  publication. Final cancellation is outside cache lookup and publication.
  A source change or cancellation during any sequence or root pass remains a
  typed failure.
* Reservation tokens are acquired before vector or scratch declaration and
  released only after the associated buffer is dropped. This ordering applies
  on success, formula-error publication, generated numerical errors,
  nonconvergence, cancellation, and every `?` exit. Numeric scratch, sequence
  metadata, and output arrays all consume the same bounded evaluator storage
  budget; none is exempt because it is “temporary”.
* No financial function consults an ambient clock, locale, filesystem, network,
  workbook recalculation service, or external rate provider.

## Demand-cache requirements

The value/matrix planner must classify financial reducers before projected
branch cache lookup. Complete sequence descriptors stay complete through a
projected lazy `IF`; a rectangular sequence must not be implicitly intersected
once per output coordinate. Scalar rate, payment, guess, control, investment,
and reinvestment arguments retain their current projected coordinate.

Direct complete references and stable literal arrays may be cacheable as
complete inputs. A computed sequence expression is cacheable only when its
complete descriptor/value, source identity, and argument order are invariant
under the projection. The conservative default is to decline the cache for a
computed sequence whose shape or contents can vary by coordinate. `MUNIT` and
other position-sensitive scalar children remain excluded from complete
argument propagation. NPV's cache key must preserve the ordered list of all
sequence-list arguments and all ReferenceList area occurrences.

Only a final finite numeric result or a stable formula-error payload may be
cached. Borrowed text, partial reducer state, a retained cell error before the
scan completes, typed evaluator failures, and a result produced before the
source/cancellation fences are never cache entries. A cache hit still checks
the outer source and final-cancellation fences.

## Open decisions that block implementation acceptance

The following are semantic decisions in the contract, rather than resource
permission to choose a host behavior:

1. Pin the IRR/RATE/XIRR numerical profile: root algorithm, tolerance,
   maximum iterations, per-term and total work charges, no-root result,
   multiple-root selection, behavior when a candidate crosses `Rate <= -1`,
   and the exact formula error kind for nonconvergence. Guess must remain an
   initial estimate, not a budget override.
2. Define the cross-argument traversal and error precedence for XIRR and XNPV
   values versus dates, including whether pairing is lockstep or one argument
   is fully scanned first. Define whether a scalar Guess/Rate formula Error is
   observed after a sequence Error and how a late typed failure supersedes it.
3. Pin `MIRR(Array Values)` admission: rectangular Reference, inline Array,
   Logical elements, ReferenceList, and computed arrays. Text/Empty omission is
   normative, but the other conversions and any zero-read shape refusal are
   not yet stated.
4. Preserve or resolve the textual distinction in `CUMPRINC`: it omits the
   positivity and period constraints printed for `CUMIPMT`, although its
   equation delegates to PPMT. The error and work behavior for a huge or
   reversed period interval must follow the chosen rule.
5. Decide whether XIRR's “first cash flow is the investment” is a hard sign
   constraint or explanatory convention, and whether XNPV's negative initial
   and positive later cash-flow language is enforced. Record the corresponding
   formula error category.
6. Assign exact formula error categories for violated constraints, invalid
   numeric domains, zero denominators, non-finite intermediate values, and
   sequence size/sign failures. Resource and provider failures must remain
   typed independently of that mapping.
7. State whether sequence conversion errors are retained while the remainder
   of a paired sequence is scanned in every function, including XNPV's
   explicit “every element is Number” rule. This is necessary to make the
   first formula Error and late typed-provider tests deterministic.

## Required focused evidence before support

The implementation review should include, at minimum:

* one-pass read traces for FVSCHEDULE, NPV across multiple ReferenceList
  areas, and XNPV; no materialized cell/text vector for those paths;
* exact read/work/cancellation limits, borrowed Text admission, formula Error
  followed by late provider/read/source/cancellation failure, and zero-read
  known ReferenceList/type refusals;
* bounded IRR/XIRR scratch at empty, exact-limit, one-over-limit, allocation,
  repeated-capacity, formula-error, cancellation, and typed-provider exits;
* iteration work and cancellation cases for IRR, RATE, and XIRR, including
  no-root, multiple-root, negative-rate, guess, overflow, crossing `-1`, and
  nonconvergence profiles;
* projected lazy-IF cases proving complete sequence references and ordered
  ReferenceLists are stable while scalar controls remain position-sensitive;
  nested MUNIT or equivalent position-sensitive scalar children must not be
  flattened or cached across coordinates;
* XIRR/XNPV unequal-size, date-order, first-date, sign, formula-error, and
  cross-argument typed-failure cases; and
* all open MIRR, CUMPRINC, constraint-error, and precision decisions encoded
  in focused semantic tests before any production-support claim.

This review records a bounded implementation plan and actionable open
decisions. It makes no production support, numerical conformance, cache
correctness, timing, allocation-count, RSS, or throughput claim.
