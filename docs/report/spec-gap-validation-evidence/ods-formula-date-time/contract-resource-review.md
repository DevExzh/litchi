# ODS date and time contract resource review

This is an independent resource, ownership, and demand-cache review of the
date/time contract in [`contract.md`](contract.md). It covers the 24 functions
listed by that contract and was performed without production, test, or Cargo
changes. The review applies ADR 0005's finite work, storage, cancellation, and
source rules and ADR 0006's typed-failure and publication rules.

The resource review disposition for the frozen handoff is **pass**. The
revised contract resolves the earlier API contradictions and states a usable
bounded implementation boundary. The final isolated validation receipts are
bound below.

## Reviewed contract identity

| Input | SHA-256 |
| --- | --- |
| `ods-formula-date-time/contract.md` | `cc77d41f487993b3438f817dd62a359ba4ec4b3ca2359a893b31aecc79bc2c7f` |

## Frozen source identity

The frozen source map is [`gates/freeze.json`](gates/freeze.json), candidate
commit `d16039ce48cb441c35461318c8634a49ae0b2437`, with 103 selected inputs.
The freeze manifest SHA-256 is
`461a76708e36b2a716cd622df45f014e009186223bd1db511c3e0ff0f7fa3561`.
The isolated gate checkout uses the retained `Cargo.lock` with SHA-256
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`.
The ambient root lock is `aa945c79965460e74a64063e1c45396eaae7a68e71072730e2bc426cded22f02`;
the gate copy is the frozen validation input.

Principal frozen implementation, integration, contract, and focused-test
hashes are:

| Frozen file | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `6db733541c9ab49e53628aab8e0fc823af8411b3a74ae47f88c50ba977446a1b` |
| `crates/litchi-ods/src/codec/formula/evaluation/date_time.rs` | `29f1495a55aac7effee9f6509deef38e5d5f747ba2b87c65a188d8759c1265e3` |
| `crates/litchi-ods/src/codec/formula/evaluation/date_time/kernel.rs` | `2bc377db1a01e28dbc4940c99c61fb43a308919b315bfc37a92ce86a67e5d0cc` |
| `crates/litchi-ods/src/codec/formula/evaluation/date_time/parser.rs` | `e71f8cb4075f4815a8c8a43ba5de48e0f5a411143e47c575c4a6de8bc4ac7149` |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `e2f9e5511e27e4a165dbeb7b0fba4ea8721231c610773a98f54b846c232c61f3` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/date_time.rs` | `1da16aae011dd8c94fafc39fdf81cc5d9f8019024ab3860bdc00b4d28962e159` |
| `crates/litchi-ods/tests/ods_formula_date_time_evaluation.rs` | `d9644432fd5dcac9272dcd3f942e4b669865db22fb6fe866a706ebd05bf04ea1` |
| `crates/litchi-ods/tests/ods_formula_date_time_limits.rs` | `fbf6ba638d7b80ebcda36d1091641ef45e42be07f9bc20a987583a3987fed411` |
| `crates/litchi-ods/tests/ods_formula_date_time_oracle.rs` | `6ab4f5845c58aa88b5890abc766229a4970409cfe581b924228bb4b595e97a05` |
| `docs/report/spec-gap-validation-evidence/ods-formula-date-time/contract.md` | `cc77d41f487993b3438f817dd62a359ba4ec4b3ca2359a893b31aecc79bc2c7f` |

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
  serial. The numeric fallback includes both simple fractions such as `1/4`
  and mixed fractions.
* Reservation-before-buffer declaration and explicit reservation-before-lease
  release order are stated for local and struct-owned buffers.

## Current foundation and value-adapter review

The calendar and timestamp foundation reviewed against the current working
tree is fixed-size and checked: civil dates carry three integers, serial
conversion uses checked bounds, and `CalculationTimestamp` stores an exact
validated `f64` bit pattern without a heap allocation. These pieces do not
read a resolver or access a host clock. The inspection VALUE path already
charges borrowed text bytes, reserves long grouped-input scratch before
allocation, drops scratch before its reservation, and fences cancellation
after release.

The current value-adapter handoff closes the earlier implementation gates:

* The scalar date/time parser adapter now uses the dedicated date/time parser
  entry points, reserves long grouped-input scratch, drops scratch before its
  reservation, and fences cancellation after release. The value-side
  `scalar_date` now routes through that charged adapter as well. Matrix
  date/time broadcasting uses a fixed four-slot argument stack, so it does not
  recreate an unreserved argument vector for every output cell.
* `date_time/kernel.rs::actual_actual` uses checked constant-time leap/year
  arithmetic for YEARFRAC basis 1. `value/date_time.rs` precharges holiday
  normalization at `O(n log n)+O(n)`, and the date-step charge includes the
  logarithmic holiday-membership bound for `binary_search`. These charges are
  paired with the scalar work/cancellation checks.
* `ConvertedResult` distinguishes an actual formula Error from a generated
  conversion error. `effective_error` now selects the earliest actual formula
  Error across scalar Date/Offset arguments and complete sequences before it
  considers generated conversion or malformed-sequence errors.
* The known `Areas(is_list)` preflight preserves an earlier materialized
  formula Error without reading the list. Computed values still follow the
  ordinary evaluation path, where reads needed to establish their type are
  legitimate.

The source is bound to the frozen manifest above, and the isolated gate
receipts below confirm that the selected source remained stable throughout
validation.

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

## Validation evidence

The frozen focused handoff currently supplies:

| Target | Result |
| --- | --- |
| `ods_formula_date_time_evaluation` | 22/22 |
| `ods_formula_date_time_limits` | 11/11 |
| `ods_formula_date_time_oracle` | 1/1 test, 132 contract vectors |

The focused cases cover direct cuboid references, rejected ReferenceLists with
zero resolver reads, row-major sequence precedence, inline and direct
Workdays admission, duplicate and fractional holidays, retained formula
Errors followed by typed provider failures, cancellation/read/storage/work
limits, borrowed text, zero and fractional WORKDAY offsets, all-off workweeks,
projected lazy IF branches, nested MUNIT scalar positions, timestamp fencing,
source changes, final cancellation, and the generated-versus-actual formula
error precedence regression.

The final custody-corrected isolated seven-gate run completed with 7/7 gates at exit 0. The
package gate ran 1,745 tests with no failures or ignored tests. The receipt
hashes are:

| Gate | Log SHA-256 |
| --- | --- |
| `ods-tests` | `026363768e8ef5a9a289523f29357ead967bcfeaa015d14f2cd17ca2b36149bb` |
| `clippy` | `a10a8c897c98b2b5da0baaae0dd26f30d0973e95ad5b076cc9d034e3b5792194` |
| `rustdoc` | `3745d757a8ba5e41843d6e62b12c97299674435efa21e178f6a49b0a14354e01` |
| `format` | `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` |
| `batch-format` | `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` |
| `boundaries` | `cdea9514ba7c77c160a8e73b2eeaf093accc70aefd1c3591af13d73e8b871fa` |
| `diff-check` | `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` |

The aggregate receipt `gates/results.json` has SHA-256
`fa5dfab0bef43d1403b22204185ae7e2e8fffe31e753c07860fb4768c228286a`.
The `source-before.json` and `source-after.json` receipts each have SHA-256
`864ccd3b76cc866e23c7db2cbdc1111702984e225cb7b630adc0befc2bde6a1e` and
contain identical selected-source maps. The verification receipt
`gates/verification.json` has SHA-256
`bdeb941ccf550e43a341df081dcd54e1ba9e70779744ecb80e785cca71739243` and
records `all_required_checks_passed: true` and `stable_sources: true`.
These receipts bind the frozen source and focused semantic/resource evidence;
they make no performance claim.

The implementation evidence should also cover:

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

Performance acceptance remains outside this resource review and requires its
own receipt; no timing, allocation, RSS, or throughput claim is made.
