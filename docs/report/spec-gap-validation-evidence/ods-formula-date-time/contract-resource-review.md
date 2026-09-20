# ODS date and time contract resource review

This is an independent resource, ownership, and demand-cache review of the
date/time contract in [`contract.md`](contract.md). It covers the 24 functions
listed by that contract and was performed without production, test, or Cargo
changes. The review applies ADR 0005's finite work, storage, cancellation, and
source rules and ADR 0006's typed-failure and publication rules.

The contract review disposition is **pass with implementation gates**. The
revised contract resolves the earlier API contradictions and states a usable
bounded implementation boundary. It does not claim production support or
implementation evidence.

## Reviewed contract identity

| Input | SHA-256 |
| --- | --- |
| `ods-formula-date-time/contract.md` | `3d70836316ccbd671a37714b6f3f773899e3dec9d1cef81d7ba50e8f3ed75413` |

## Resolved contract findings

The revised contract now explicitly resolves the earlier review findings:

* WORKDAY truncates a finite fractional Offset toward zero, rejects non-finite
  or checked-domain overflow, preserves the exact input serial for zero Offset,
  and reattaches the input fraction after nonzero date stepping.
* DateSequence rejects a ReferenceList before resolver reads. A single
  cuboid Reference remains a sequence source and is streamed in descriptor
  sheet order and the selected row-major order. LogicalSequence keeps the
  distinguished-Logical rule for direct references: only Logical and Error
  cells are admitted, while Empty and Text are skipped.
* Inline LogicalSequence arrays define Number, Logical, Empty, and Text
  behavior and require seven positions. Inline DateSequence elements use the
  finite Number conversion and retain out-of-domain numeric errors while the
  rest of the array is scanned.
* Formula Errors are retained while an admitted sequence is consumed. Later
  resolver, source, cancellation, allocation, or resource failures remain
  typed and supersede the retained formula result. Optional sequences are
  consumed even for zero Offset and all-off workweeks.
* Holiday dates are floored and normalized to bounded integer days, duplicate valid dates
  are harmless, and the contract requires fallible retained-day charging
  without an unbounded range-sized raw-cell vector.
* The timestamp profile bounds no-Year EASTERSUNDAY, including the unsupported
  9956-after-Easter and 9957–9999 cases. Timestamp-backed cache entries are
  restricted to the same explicit snapshot.
* Date-only extraction is explicitly floor-based for the supported negative
  serial extension, while Integer and WORKDAY Offset remain toward-zero
  conversions. DateParam parser fallback is numeric-only, so sharing VALUE
  helpers cannot accidentally widen or reinterpret date/time text.
* The YM rule is now executable: compute the signed month difference, adjust
  for the day boundary, and apply Euclidean modulo only after that adjustment.
  The DATEVALUE/TIMEVALUE fallback is likewise explicit: it uses the numeric
  VALUE branch only, with date/time dispatch disabled; DATEVALUE floors a
  finite fallback serial and TIMEVALUE preserves it as the raw finite time
  serial.
* Reservation-before-buffer declaration and explicit reservation-before-lease
  release order are stated for local and struct-owned buffers.

## Incremental foundation review (source still in progress)

The calendar and timestamp foundation reviewed against the current working
tree is fixed-size and checked: civil dates carry three integers, serial
conversion uses checked bounds, and `CalculationTimestamp` stores an exact
validated `f64` bit pattern without a heap allocation. These pieces do not
read a resolver or access a host clock. The inspection VALUE path already
charges borrowed text bytes, reserves long grouped-input scratch before
allocation, drops scratch before its reservation, and fences cancellation
after release.

The following implementation gates remain open while the date/time value
adapter is being integrated:

* `evaluation/date_time/parser.rs::has_month_name` currently calls
  `text.to_ascii_lowercase()` once per entry in a 24-entry month table. That
  allocates temporary owned strings in the date parser and has no reservation
  or allocation-failure mapping. The probe must become allocation-free (for
  example, a borrowed ASCII case-insensitive scan) or use explicitly charged,
  fallible scratch. This is a concrete resource blocker for the adapter.
* The same adapter currently delegates to the generic `parse_value` entry
  point. The integrated DATEVALUE/DateParam and TIMEVALUE/TimeParam paths
  must call the dedicated numeric-fallback parsers, with their caller-owned
  scratch reservation and post-release cancellation fence. Delegating the
  full VALUE grammar would re-enable date/time dispatch after the contract's
  numeric-only fallback gate and can also bypass the long grouped-input
  reservation path.
* `date_time/kernel.rs::actual_actual` scans the inclusive year range (up to
  the 1..=9999 profile) twice for YEARFRAC basis 1. The bounded range is safe
  for memory, but an unchecked scalar loop can bypass a low execution-work
  budget or cancellation. The final adapter must either use checked constant
  time leap/year arithmetic or charge and checkpoint this work at an agreed
  cadence.

These are review gates rather than a source disposition: the date/time module
is not yet a frozen, buildable integration in the current shared tree, so this
document makes no production-support claim.

## Resource and ownership requirements for implementation

Reference-backed DateSequence and LogicalSequence inputs must retain only
bounded descriptors and stream cells in the chosen sheet/row/column order.
Every cell that must be inspected, including a cell later skipped because it
is Empty or Text, needs cell work charging, reference-cell-limit checking, and
cancellation checks before and after the resolver call. A successful read is
counted only through the existing resolver read path. ReferenceList and
scalar-shape refusals that are known from the AST/runtime descriptor must
return before resolver reads; evaluating an arbitrary computed expression may
legitimately perform reads before its resulting type is known.

Inline arrays need bounded element traversal and work charging as well as
their existing array reservation. The implementation should pin the exact
profile for inline DateSequence Empty, Text, Logical, and Error elements by
tests; “same finite Number conversion” must use the date family's explicit
domain and must not silently widen the existing VALUE domain.

Holiday storage should use a checked bounded representation, such as a
normalized day vector sorted and deduplicated after the complete scan. It must
reserve before growth, charge normalization and membership work, and retain a
fixed-size formula/generated-error record rather than storing formula-error
cells or borrowed text. Duplicate valid dates do not multiply membership, but
duplicate formula Errors still participate in formula-error precedence. An
out-of-domain numeric holiday retains `#NUM!` and does not stop the scan.

The exact-seven Workdays check needs two explicit paths: inline arrays count
their seven traversed positions, while direct references count admitted
Logical/Error elements after Empty/Text/Number filtering. Once sequence
evaluation has begun, a generated malformed-sequence result must not stop the
scan if a later provider, cancellation, allocation, or resource failure can
supersede it. A retained formula Error has precedence over an ordinary
generated formula result according to the contract's source order.

NETWORKDAYS and WORKDAY must charge date-iteration work even when no cell
references are present. Seven-day-cycle optimizations may reduce loop count,
but the charged work must remain proportional to the skipped day span or
candidate steps, with checked arithmetic and cancellation checkpoints. A
finite Offset whose candidate date is outside the profile is `#NUM!`; a valid
operation that exhausts the execution work budget remains a typed resource
failure. All-off workweeks and zero Offset still consume supplied sequence
arguments before publishing their ordinary result or `#NUM!` refusal.

Declare holiday vectors, parser scratch, sequence metadata, matrix output, and
their reservation tokens in the documented order. Local buffers must be
declared after their reservation token, struct buffers before their reservation
field, and every reservation must be released before its enclosing budget
lease. Resolver text remains borrowed through skip/conversion paths; no range
or holiday scan may clone a cell's complete text merely to test admission.

## Matrix and demand-cache requirements

Date and Offset arguments remain ordinary scalar intersection/projected
arguments. NETWORKDAYS and WORKDAY consume Holidays and Workdays completely
for each projected scalar result; their references must never be implicitly
intersected or partially selected. The sequence descriptors themselves are
not demand-cache payloads, so any reuse of a complete reducer result must be
limited to invariant scalar arguments and a stable source/shape context.

NOW, TODAY, and no-Year EASTERSUNDAY are invariant under output projection
within one explicit `CalculationTimestamp`. A cache shared beyond one
evaluator must include timestamp identity, source version, projected-shape,
position, and cancellation identity; an evaluator-local cache must still
validate the timestamp capability before a hit. A missing timestamp is typed
`Unsupported(CalculationClock)` and must never be cached as a formula Error.

Date scalar parameters and Offset must preserve projected position for nested
position-sensitive expressions. In particular, the date integration must keep
the existing MUNIT scalar-size exclusion: a nested MUNIT in a Date, Offset,
or other scalar slot cannot be flattened by complete-sequence propagation.
MUNIT used to construct a complete inline sequence may be consumed as an
array, but its scalar size argument remains position-sensitive and its typed
provider failures must not be hidden by shape or cache probes.

Source-version and final-cancellation fences remain outside cache lookup and
publication. Typed failures are never converted to formula Errors, cached, or
caught by IFERROR/IFNA; a matrix result is published only after every selected
cell and sequence has completed successfully at the typed level.

## Focused validation required before an implementation disposition

The implementation evidence should cover:

* direct cuboid references, rejected ReferenceLists with zero resolver reads,
  row-major/sheet-order error precedence, and computed IF reference results;
* numeric, Logical, Empty, Text, and Error inline Workdays values; direct
  reference filtering; malformed admitted cardinalities; and the empty third
  slot used to reach the fourth argument;
* duplicate and fractional-date holidays, out-of-domain holiday numbers,
  formula Error followed by a provider failure, cancellation, read limits,
  storage limits, and borrowed-text admission;
* zero and fractional WORKDAY offsets, all-off workweeks, reversed and
  maximum-domain NETWORKDAYS intervals, large-offset overflow, work-budget
  refusal, and cancellation during optimized date stepping; and
* projected lazy IF cases with complete holiday/workweek references, nested
  MUNIT in scalar date/offset positions, timestamp changes across evaluator
  contexts, timestamp absence, source changes, and final cancellation.

No source or test PASS is claimed by this document. The resource contract is
ready for implementation once these gates preserve the stated typed-failure,
work-charge, cache-identity, and ownership rules.
