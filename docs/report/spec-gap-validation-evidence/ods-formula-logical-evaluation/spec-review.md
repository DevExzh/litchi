# ODS OpenFormula logical-function review

Status: normative review of OpenDocument 1.4 Part 4 logical functions, plus a
read-only review of the current scalar-evaluator candidate. The candidate is a
bounded scalar foundation; this report does not claim full OpenFormula
evaluation, array evaluation, reference resolution, or recalculation.

## Checked source and scope

The normative source is the checked-in `3rdparty/specs/` artifact:

* `3rdparty/specs/OpenDocument-v1.4-os.zip`, SHA-256
  `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`.
* Entry `part4-formula/OpenDocument-v1.4-os-part4-formula.html`, SHA-256
  `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.

The primary logical-family anchors are `a_6_15_1_General` through
`a_6_15_10_XOR`. The evaluation and type rules used below are
`a_3_2_3_Operator_and_Function_Evaluation`, `a_4_5_Logical__Boolean_`,
`a_4_6_Error`, `a_4_7_Empty_Cell`, `a_6_1_General`,
`a_6_2_Common_Template_for_Functions_and_Operators`,
`a_6_3_5_Conversion_to_Number`, `a_6_3_7_Conversion_to_NumberSequence`,
`a_6_3_8_Conversion_to_NumberSequenceList`, and
`a_6_3_12_Conversion_to_Logical`.

The review is limited to the nine functions in §6.15 and their scalar
argument/evaluation boundary. Section 3.3's range/array aggregation remains a
future capability: a scalar operation may refuse a reference or array, but it
must not silently select an element or flatten a range.

The design is also bounded by accepted ADRs 0002, 0004, 0005, 0006, 0023,
and 0024. In particular, formula errors are values, cancellation and resource
exhaustion are operation failures, external resolution is an explicit
capability, and evaluation does not mutate a workbook, cache, package, or
network resource.

The final reviewed candidate digest is `SHA-256
cb47a3c6a6a0169f225d0e4b979fe5ec9ebf629bc1a5fdef86dca18f9c3c8efa` for
`crates/litchi-ods/src/codec/formula/evaluation.rs`.
The focused source is `bea3a987b1a5d922d6ab83e317b8ee997ae3a187685d246b2b7ef3735bb89fad`.
It includes the corrected `IF(FALSE();1)` assertion as Logical FALSE,
matching the omitted-false default in §6.15.4, and missing-operand handler
fixtures (`IFERROR(;2)`, `IFNA(;2)`, and selected missing alternatives). The
final focused suite has 15 integration groups. The parent gate receipt reports
894 tests across 53 targets, two doctests, Clippy, rustdoc, and formatting
passing for this final source vector.

## Complete §6.15 family

| Function | Normative signature and arity | Result | Evaluation and error rule |
| --- | --- | --- | --- |
| `AND` | `AND({ Logical\|NumberSequenceList L }+)`; one or more parameters; zero-parameter behavior may be Logical or Error | Logical | All TRUE gives TRUE, any FALSE gives FALSE. The parameter is converted through `NumberSequenceList`, so a scalar Number, Text, or Logical follows Conversion to Number. General Error propagation applies. In array context it aggregates all arguments and does not return a matrix. |
| `FALSE` | `FALSE()`; exactly zero parameters | Logical (possibly represented as Number) | Constant FALSE. |
| `IF` | `IF(Logical Condition [; [Any IfTrue] [; [Any IfFalse]]])`; one through three supplied slots | Any | Evaluate the condition, then evaluate only the selected branch. The other branch is never evaluated. Explicit empty slots and omitted optional parameters have different defaults (see below). |
| `IFERROR` | `IFERROR(Any X; Any Alternative)`; exactly two parameters | Any | Compute `X` once. Return `Alternative` only when `X` is a formula Error; otherwise return `X`. The alternative is therefore lazy under the `IF` equivalence. |
| `IFNA` | `IFNA(Any X; Any Alternative)`; exactly two parameters | Any | Compute `X` once. Return `Alternative` only for the `#N/A` Error; other formula Errors pass through and the alternative is not evaluated. |
| `NOT` | `NOT(Logical L)`; exactly one parameter | Logical | Convert the parameter to Logical and invert it. |
| `OR` | `OR({ Logical\|NumberSequenceList L }+)`; one or more parameters; zero-parameter behavior may be Logical or Error | Logical | All FALSE gives FALSE, any TRUE gives TRUE. It has the same NumberSequenceList and Error rules as `AND`; in array context it aggregates rather than iterates elementwise. |
| `TRUE` | `TRUE()`; exactly zero parameters | Logical (possibly represented as Number) | Constant TRUE. A distinct Logical need not compare equal to Number 1, but converts to Number 1 when Number is required. |
| `XOR` | `XOR({ Logical L }+)`; one or more parameters | Logical | Logical addition modulo 2: an odd TRUE count gives TRUE, an even count gives FALSE. Each parameter uses Conversion to Logical. |

Function names are case-insensitive (§6.1). The family is logical, not
bitwise: for example, `AND(12;10)` returns TRUE rather than the bitwise value
8 (§6.15.1).

The general call algorithm in §3.2.3 is eager: compute all argument values,
propagate an Error when the function does not suppress it, perform the
specified implicit conversions, and call the function. `IF`, `IFERROR`, and
`IFNA` are the defined control-flow exceptions. `AND`, `OR`, `XOR`, `NOT`,
`TRUE`, and `FALSE` therefore visit their supplied children even when a
partial result is already decisive or the call will fail an arity constraint.
An evaluator may short-circuit on a formula Error when the function does not
suppress it, but the scalar profile deliberately keeps aggregate Boolean
evaluation eager so later Error values, unsupported capabilities, work limits,
and cancellation remain observable.

Unless a function says otherwise, §6.1 recommends returning the leftmost
source-order Error when more than one Error is provided. Values can be formula
Errors (`#N/A`, `#DIV/0!`, `#VALUE!`, and so on); evaluator failures such as
`Unsupported`, `Cancelled`, `ResourceLimit`, or `Allocation` are not formula
Errors and must not be caught by `IFERROR` or `IFNA`.

## `IF` slots and defaults

The seven forms in §6.15.4 are distinct:

| Source form | IfTrue when selected | IfFalse when selected |
| --- | --- | --- |
| `IF(Condition)` | Logical TRUE | Logical FALSE |
| `IF(Condition;)` | Number 0 | Logical FALSE |
| `IF(Condition;IfTrue)` | supplied `IfTrue` | Logical FALSE |
| `IF(Condition;;)` | Number 0 | Number 0 |
| `IF(Condition;;IfFalse)` | Number 0 | supplied `IfFalse` |
| `IF(Condition;IfTrue;)` | supplied `IfTrue` | Number 0 |
| `IF(Condition;IfTrue;IfFalse)` | supplied `IfTrue` | supplied `IfFalse` |

An omitted optional parameter is different from a semicolon-delimited empty
slot. The condition is required; a missing condition has no default. Outside
these seven `IF` forms, the scalar profile returns a formula `#VALUE!` for an
unsupported arity rather than silently extending the function. §6.2 permits
documented evaluator extensions, but an extension would be a separate
compatibility profile.

`IFERROR` and `IFNA` have no optional defaults. A missing required operand is a
formula error at the scalar boundary; it is not a fabricated Empty, zero, or
false value. An alternative that is actually selected may itself return an
Error, which is then returned as that function's result.

## Conversion choices for this scalar profile

OpenFormula permits implementation-defined text conversion in several places,
so the profile must state its choice rather than imply a universal ODF rule.

* `Logical` is represented distinctly from `Number`. Conversion to Logical
  maps Number zero to FALSE and every admitted nonzero finite Number to TRUE;
  Logical passes through. Direct Text-to-Logical conversion returns formula
  `#VALUE!` in this locale-independent profile. ODF §6.3.12 also permits FALSE
  or a case-insensitive locale-dependent attempt, so this is a documented
  profile choice, not a claim that all evaluators must reject text.
* `AND` and `OR` have the separate `NumberSequenceList` alternative in their
  signatures. For direct scalar Text, §6.3.7 and §6.3.8 explicitly route
  through Conversion to Number. Consequently this profile accepts
  `AND("1")` as TRUE and `OR("0")` as FALSE, using the existing
  locale-independent decimal parser. Invalid numeric text produces formula
  `#VALUE!`; it is not treated as a Logical spelling. A future reference
  resolver must preserve the sequence rule that referenced Text and Empty
  cells are excluded from a NumberSequence, while direct scalar Text is
  converted.
* `NOT`, `IF` conditions, and `XOR` expect Logical and therefore use the
  direct Logical conversion policy. `IF("1";1;2)`, `NOT("TRUE")`, and
  `XOR("1")` return `#VALUE!` in this profile.
* `TRUE` and `FALSE` remain Logical values. Numeric operators convert them to
  1 and 0, while equality with Number 1 follows the distinct-Logical choice.
* Formula Error values pass through ordinary conversion unchanged unless an
  explicitly defined handler catches them. Empty is a separate ODF type; the
  current scalar cut has no reference resolver and does not globally map an
  Empty to zero, false, or empty Text.

## Required behavior fixtures

These examples are a compact normative/profile matrix. `Error` in the result
column means a formula-level scalar Error; `Unsupported(Reference)` is a Rust
capability refusal and is intentionally a different result class.

| Expression | Expected scalar result or boundary |
| --- | --- |
| `TRUE()` / `FALSE()` | Logical TRUE / Logical FALSE |
| `NOT(0)` / `NOT(-2)` | Logical TRUE / Logical FALSE |
| `AND(1;0)` / `OR(0;2)` | Logical FALSE / Logical TRUE |
| `AND("1")` / `OR("0")` | Logical TRUE / Logical FALSE under the profile's NumberSequenceList conversion |
| `XOR(1;1)` / `XOR(1;1;1)` | Logical FALSE / Logical TRUE |
| `IF(1;2;3)` / `IF(0;2;3)` | Number 2 / Number 3 |
| `IF(1;)` / `IF(0;)` | Number 0 / Logical FALSE |
| `IF(1;;3)` / `IF(0;;3)` | Number 0 / Number 3 |
| `IF(1;2;)` / `IF(0;2;)` | Number 2 / Number 0 |
| `IFERROR(#N/A;5)` | Number 5; the formula Error is suppressed |
| `IFERROR(7;#N/A)` | Number 7; the alternative is not evaluated |
| `IFNA(#N/A;5)` | Number 5 |
| `IFNA(#DIV/0!;5)` | formula `#DIV/0!`; only `#N/A` is caught |
| `IF(1;5;[.A1])` | Number 5; the unselected reference is not resolved |
| `IFERROR([.A1];5)` | `Unsupported(Reference)`; evaluator failures are not caught |
| `AND(FALSE();[.A1])` / `OR(TRUE();[.A1])` | `Unsupported(Reference)` because aggregation is eager |
| `AND(#N/A;TRUE())` | formula `#N/A`, retaining the leftmost Error |
| `AND()` / `OR()` | profile may return a Logical identity or formula Error; this profile uses formula `#VALUE!` |
| `XOR()` / `NOT()` | formula `#VALUE!` for the unsatisfied one-or-more/exactly-one constraint |

## Candidate implementation disposition

The current read-only candidate in
`crates/litchi-ods/src/codec/formula/evaluation.rs` follows the contract above:

* `visit_function` identifies all nine names and routes control flow through
  an explicit `Apply` dispatch frame. A selected branch is pushed
  only after its condition or handled value is available, so an unselected
  reference, unsupported function, long expression, or owned text is neither
  visited nor allocated. `IFERROR` and `IFNA` distinguish formula Error values
  from Rust failures.
* `AND` and `OR` use the `to_number_sequence` path, which delegates to the
  existing Number conversion and therefore preserves the required numeric-text
  behavior. `NOT`, `IF`, and `XOR` use `to_logical`. Aggregate application
  retains the first source-order Error while still consuming all eagerly
  evaluated values.
* Explicit frame/value stacks avoid recursive evaluator traversal. Child-count
  work, function-name bytes, conversion bytes, and text work are charged;
  stack and owned-text growth is fallibly reserved. `schedule_eager_function`
  checks cancellation while pushing reverse-order child frames with a
  processed-item counter, and each frame checks again before execution. The
  selected output reservation remains attached to `EvaluatedScalar` until the
  result is dropped.
* The implementation returns `#VALUE!` for the permitted `AND()`/`OR()` zero
  case and for unsupported logical arities. These are valid profile choices
  under the specification's constraint language. It keeps references, arrays,
  names, labels, and other function families as typed refusals.

No additional semantic blocker was found in this bounded scope. The candidate
review is read-only and does not replace the parent gates; any later source
change requires refreshing the source digest and rechecking the listed lazy,
eager, conversion, and error fixtures. The evaluator's indivisible finite
numeric parser and the unresolved reference/array/host-locale semantics remain
explicit boundaries from the earlier scalar review.
