# ODS OpenFormula scalar evaluation review

Status: normative and API design review. No evaluator implementation is
approved by this document. The existing ODS formula trees remain inert and the
evaluation boundary must be added as a separately bounded, read-only operation.

## Scope and primary sources

This review covers the scalar part of OpenFormula 1.4: the value model in Part
4 chapter 4, expression calculation in §§3.2.2–3.2.3, host properties in §3.4,
numeric constraints in §§3.6–3.7, implicit conversions in §6.3, and the
standard operators in §6.4. It does not claim array evaluation, the function
catalog, recalculation scheduling, or external-data activation.

The checked-in normative source is:

* `3rdparty/specs/OpenDocument-v1.4-os.zip`, SHA-256
  `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`.
* `part4-formula/OpenDocument-v1.4-os-part4-formula.html` inside that ZIP,
  SHA-256 `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.

The relevant HTML anchors are `a_3_2_Expression_Evaluation`,
`a_3_3_Non-Scalar_Evaluation__aka_'Array_expressions'_`,
`a_3_4_Host-Defined_Behaviors`, `a_3_6_Numerical_Models`,
`a_3_7_Basic_Limits`, `a_4_Types`, `a_6_3_Implicit_Conversion_Operators`,
and `a_6_4_Standard_Operators`, with the section-specific anchors cited
below. The review was made at repository HEAD
`9072e969e34263be73d5b8d2eb2700478d182101`.

## Normative scalar contract

The evaluator must preserve the distinction between these values:

* `Number`, including fractional values. Integer-only arithmetic cannot claim
  the `/` operator. Date, time, date-time, percentage, and currency are Number
  subtypes; the specification does not require tracking their subtype, and
  subtype results after arithmetic are otherwise often unspecified.
* `Text`, including the empty string. String indexes are one-based. Unicode
  normalization of formulas, results, and inputs is implementation-defined;
  an implementation must document its choice and must not silently normalize
  operands under an unrelated policy.
* `Logical`, which may be distinct from Number or may be represented by 1 and
  0. The representation choice changes `1=TRUE()` and number-sequence
  inclusion, so it belongs in an explicit semantic policy.
* `Empty`, which is neither zero, the empty string, nor an Error (in
  particular, it is distinct from `#N/A`). A reference to an empty cell is
  handled by the conversion requested by the operation.
* `Error`, with `#N/A` distinguished from other errors. An operation or
  function receiving an Error normally returns one of its input Errors. The
  ODF cell encoding of a computed error as `office:value-type="string"` and
  `office:string-value` is storage behavior; it does not turn the formula
  value into Text.
* `Reference`, `ReferenceList`, and `Array`, which are non-scalar values. A
  multi-cell reference or an array with more than one element cannot be
  treated as a scalar without the §3.3 intersection rules. A reference list
  cannot be converted to an array and remains ordered.

This implies a value type separate from `litchi_core::sheet::CellValue`.
`CellValue` currently includes storage-oriented `Formula` wrappers, an
untyped `Error(String)`, `Int`/`Float`/`DateTime` distinctions, and no
Reference, ReferenceList, or Array value. Mapping an ODS Error to a string or
an Empty value to a zero at this boundary loses required semantics.

Expression calculation follows §3.2.2: constants return their value;
operators use the §5.5 precedence and associativity table; function calls and
named expressions use their definitions; quoted labels, automatic
intersection, and arrays return their value. After calculation, a Reference is
converted only when a single non-reference value is needed, and an Array is
projected for the display area under §3.3. The scalar cut should therefore
return an explicit `UnsupportedArrayContext` or `UnsupportedReference` when
the caller did not supply the corresponding context. It must not select the
first array cell, flatten a range, or stringify a reference.

Under §3.2.3, expression arguments are eager unless a function specification
allows otherwise (for example, `IF`); an Error argument may short-circuit when
the function does not suppress propagation; wrong types invoke the relevant
implicit conversion; and conversion failure produces an Error. A registry for
future functions must carry a lazy/eager argument policy, rather than making
all calls eager in the evaluator core. For deterministic, side-effect-free
evaluation, child values can be visited in source order and the first
propagated input Error retained; this is an implementation ordering choice and
must not make resolution or function calls observable.

## Required conversions

The following are the minimum conversion rules for the scalar operators and
the conversion hooks that future functions will use.

| Expected type | Required behavior | Host or function boundary |
| --- | --- | --- |
| Scalar | Number, Logical, and Text pass through. A single-cell reference is read. A multi-cell reference uses implied intersection from §6.3.3. | Requires the formula position and a bounded reference resolver. |
| Number | Number passes through. Logical becomes 0/1. Text conversion is implementation-defined and may be 0, an Error, or a locale-aware attempt. A reference intersects first; an empty referenced cell becomes 0. | `HOST-LOCALE` and `HOST-PRECISION-AS-SHOWN` must be explicit. Embedded Text and referenced Text may intentionally differ. |
| Integer | Convert to Number, then apply the function's specified rounding. `INT` is toward negative infinity, `ROUND` is nearest with halfway away from zero, and `TRUNC` is toward zero. | Do not apply a global cast or truncation. |
| Logical | Number zero/nonzero becomes FALSE/TRUE. Text conversion is implementation-defined and may be FALSE, an Error, or a locale-aware attempt. A referenced empty cell is FALSE. | `HOST-LOCALE`; direct Empty is not a blanket false/zero rule. |
| Text | Number converts to Text with no whitespace. Text passes through. Logical converts to uppercase `TRUE`/`FALSE`. A referenced empty cell converts to an empty string. | Number formatting must be deterministic or explicitly locale-backed. |
| DateParam/TimeParam | Number passes through. Text uses `DATEVALUE`/`TIMEVALUE`, with an implementation-defined numeric fallback. Logical behavior is implementation-defined. A referenced empty cell becomes 0. | Epoch, two-digit year, and locale are host settings. |

The direct `Empty` case is deliberately not globally normalized: §§6.3.5,
6.3.12, and 6.3.14 define the empty-cell behavior through a Reference. A
function may specify another rule, but a shared conversion helper must not
turn every Empty into all three of 0, FALSE, and `""`.

## Operator behavior

The evaluator needs the following table, in addition to the exact §6.4
constraints:

| Operator | Result and conversion | Required error/domain behavior |
| --- | --- | --- |
| `Left + Right` | Number + Number -> Number | Use Conversion to Number; preserve an input Error. |
| `Left - Right` | Number - Number -> Number | Same. |
| `Left * Right` | Number * Number -> Number | Same. |
| `Left / Right` | Number / Number -> Number | Divisor zero returns an Error. |
| `Left ^ Right` | Number `POWER` Number -> Number | `0^0` is constrained; domain/overflow failure returns an Error. |
| `Left = Right` | Scalar -> Logical | Different types are FALSE. Number equality may use a documented mantissa-scaled approximation. Text comparison obeys `HOST-CASE-SENSITIVE`; Logical comparison is identity. Errors cannot be compared to identify the same Error. |
| `Left <> Right` | Any -> Logical, defined as NOT(`=`) for non-Errors | Either Error returns an Error; it cannot distinguish two Errors. |
| ordered comparison | Scalar -> Logical | Compare Number/Number, Text/Text, or Logical/Logical (Logical via Number). For different types, TRUE/FALSE/Error is implementation-defined and must be documented. |
| `Left & Right` | Text + Text -> Text | Use Conversion to Text; preserve input Errors. |
| `Left %` | Number -> Number / 100 | Use Conversion to Number. |
| prefix `+` | Any -> Any, unchanged | It performs no conversion; a Reference may remain a Reference or return an Error. |
| prefix `-` | Number -> Number, `0 - Right` | Use Conversion to Number. |

The §5.5 precedence table is part of the evaluator contract: left-associative
`:`, `!`, and `~`; right-associative prefix `+`/`-`; postfix `%`; left-
associative `^`, `*`/`/`, `+`/`-`, `&`, then comparisons. In particular prefix
minus binds more tightly than power, and `^` is left-associative. Parentheses
override the table. The parser already retains these operator kinds, but the
evaluator must not substitute Excel precedence or Excel's comparison coercion.

## Host-dependent semantics and deferred features

`HOST-CASE-SENSITIVE`, `HOST-PRECISION-AS-SHOWN`, and `HOST-LOCALE` directly
affect this scalar cut. `HOST-NULL-DATE` defines serial zero, `HOST-NULL-YEAR`
defines two-digit years, and the implementation-defined 1900 leap-year rule
affects date conversion. These must be supplied by an explicit options value;
the evaluator must not consult process locale, display state, wall clock, or
an implicit workbook-global setting. A text comparison that uses `EXACT` in a
future function remains case-sensitive regardless of the host setting.

The reference resolver in §3.4 is an implementation-defined function. It must
be an explicit capability accepting the parsed Unicode Reference and formula
position, with bounded cell/area work. A missing resolver is a capability
refusal, not a fabricated zero or an accidental lookup in a different
workbook. External source IRIs, DDE, hyperlinks, `INDIRECT`, `INFO`, and other
external or volatile functions remain inert until a caller supplies a narrowly
scoped provider. No provider may perform network I/O as a side effect of
constant scalar evaluation.

§3.3's implicit intersection and matrix iteration are outside the first
scalar implementation. It is valid for the syntax parser to accept inline
arrays, including a ragged syntax shape; an evaluator that needs a rectangular
matrix must validate that shape at the array boundary and must not reject the
syntax parser's ragged representation by conflating syntax with evaluation.
Likewise, range (`:`), intersection (`!`), and reference concatenation (`~`)
produce reference values and require a reference context; they are not string
or scalar arithmetic operators.

## Proposed owned API and execution boundary

The public entry point should be owned by `litchi_ods` and consume the
immutable ODS expression tree. A suitable shape is:

```rust,ignore
pub fn evaluate_scalar(
    expression: &Expression,
    context: &OdsEvaluationContext<'_>,
    limits: &EvaluationLimits,
) -> Result<OdsValue, EvaluationFailure>;

pub enum OdsValue {
    Number(Number),
    Text(String),
    Logical(bool),
    Empty,
    Error(OdsError),
    Reference(Reference),
    ReferenceList(ReferenceList),
    Array(Array),
}

pub struct OdsEvaluationContext<'a> {
    pub position: FormulaPosition,
    pub host: HostOptions,
    pub references: Option<&'a dyn ReferenceResolver>,
    pub functions: &'a dyn FunctionRegistry,
    pub execution: &'a ExecutionContext,
}
```

The exact names can change, but the boundaries are required:

* `OdsValue` is a formula value, so an ODF `Error` is not a Rust failure.
  The first scalar profile may expose only the scalar variants and return a
  typed unsupported capability for references, arrays, names, and functions it
  does not implement; adding dormant variants is not required before those
  boundaries exist.
  `EvaluationFailure` is reserved for cancellation, resource exhaustion,
  malformed/inconsistent evaluator state, and a caller-denied host capability.
  `OdsError` should be a non-exhaustive typed code with at least a distinct
  `NotAvailable`/`#N/A` case, rather than a free-form diagnostic string.
* `HostOptions` explicitly carries case sensitivity, precision-as-shown,
  locale/date policies, and whether Logical is distinct from Number. Text
  numeric and text logical conversions need explicit policy variants because
  Part 4 makes them implementation-defined. Unicode normalization should be a
  visible policy; preserving input code points is the safest default.
* `ReferenceResolver` and named-expression lookup are optional capabilities,
  and must be tied to the immutable source snapshot and formula position.
  Reference reads return ODS values and are charged; source IRIs require a
  separate policy/provider. There is no ambient workbook, network transport,
  current time, or cache.
* `FunctionRegistry` owns arity, expected conversion types, error-suppression
  rules, and lazy arguments. It must identify the `OpenFormula` dialect and
  return a typed unsupported/host refusal for disabled extensions. The ODS
  scalar evaluator must not dispatch through an Excel-only name table.
* `ExecutionContext`/`EvaluationLimits` must provide finite maximum steps,
  depth, argument count, text bytes, resolved cells/areas, and result storage.
  Charge the node/operator/coercion before work, check cancellation at each
  charged step and before each resolver iteration, and use checked arithmetic
  plus fallible reservations. A first scalar cut can reject arrays and
  multi-cell references before allocating their materialized values.

Evaluation should walk the immutable flat arena with an explicit work stack or
a bounded recursive depth, stage values in local owned storage, and return the
single result only after the walk succeeds. It must never write a formula
cache, alter a cell, refresh a cached value, or publish a package. A future
recalculation feature can stage a dependency closure and feed it to the
existing transaction/commit machinery, but that is a separate operation with
its own source checks and atomic publication.

## Existing evaluation owner and compatibility blockers

`litchi-eval` is tracked and is the existing shared runtime owner, but its
current public engine cannot be used as an ODS evaluator without an explicit
dialect/kernel boundary:

* `crates/litchi-eval/src/parser/{ast,expr,literal,reference}.rs` is an
  Excel-oriented compatibility parser. Its AST lacks ODF `%`, `^`, `&`,
  reference `:`, `!`, `~`, arrays, ReferenceList, and ODS source-IRI metadata;
  its reference spelling uses Excel `!` and its function arguments use commas.
  It also accepts Rust floating-point spellings instead of the §5.3 number
  lexical rules.
* `crates/litchi-eval/src/engine.rs` and `engine/bin_op.rs` use
  `CellValue`, generic `to_number`/`to_bool`/`to_text` helpers, and
  numeric-first comparison. That makes `"1"=1` or unrelated types compare by
  coercion, maps any nonempty string to TRUE, exposes formula/error wrappers as
  text, and does not implement the ODF host policies or typed Error
  propagation. It also has no explicit finite evaluation budget/cancellation
  parameter.
* `crates/litchi-eval/src/lib.rs`'s `FormulaEvaluator` prefers embedded cached
  results and inserts evaluated values into an internal `eval_state` cache.
  Internal memoization alone need not violate immutable workbook semantics, but
  cached-result preference and stateful cache behavior must not be silently
  included in a read-only ODS scalar operation, which should evaluate the
  supplied ODS AST/value snapshot without publishing or mutating it. Its
  optional `Fetch` is an explicit capability and remains network-inert without
  a provider, yet it must not be enabled implicitly by the ODS scalar API.

The recommended ownership is therefore `litchi-ods` for the OpenFormula AST
adapter, `OdsValue`, conversion/operator semantics, ODF function registry, and
public ODS evaluation façade. Shared execution code may be extracted from
`litchi-eval` only as a runtime-neutral kernel that accepts an explicit
`OpenFormula` dialect, the ODS value/error model, an ODS AST, and the caller's
budget/cancellation context. It must not reuse `litchi-eval::parser::Expr`,
silently route ODS formulas through Excel semantics, or create a second
uncoordinated formula registry. Keeping the first implementation format-owned
also avoids introducing a dependency cycle while the shared kernel contract is
being defined.

## Conformance and acceptance boundary

Part 4 §2.3 defines small, medium, and large evaluator groups. A scalar
operator implementation alone must not advertise an OpenDocument Formula
Evaluator or Small Group Evaluator: the small group additionally requires the
basic limits, the chapter 5 syntax subset, all relevant implicit conversions,
the listed standard functions, and all operators except only reference union
(`~`), with the stated allowance for complex numbers, inline arrays, and
multi-area references. The implementation can document a bounded scalar
profile until those obligations are met.

If claiming the §3.7 basic limits later, use at least 1024 interchange
characters, 30 parameters where a prototype permits a list, 32,767 ASCII
characters in a string, and seven function nesting levels. Parser limits and
evaluation limits are separate: a large source/node budget does not imply a
safe result or reference budget.

## Source inventory

The reviewed implementation files had these SHA-256 values at the review
point:

```text
crates/litchi-ods/src/codec/formula.rs             e5681591c4c1c196e21141eadee7fd046bbb668b6defc8ab5ced1706854ee072
crates/litchi-ods/src/codec/formula/expression.rs  b75b78b75b0b1dbf660f736eadddee335c18c2121fd0bc78198c7eca1c58ef04
crates/litchi-ods/src/codec/formula/reference.rs   163e5131e0f787059641df30fda0b0cb2b7582fdf414f6263d1ebfbb80a47c40
crates/litchi-eval/src/lib.rs                      803d8a384e57cc65a2ca32ca5e2f1696ca5e412c8874cafce83703690b77507c
crates/litchi-eval/src/engine.rs                   90a032120ef7503ac71019727013dee57c89e75440647276d1d3b338664d4634
crates/litchi-eval/src/parser/ast.rs               2e44e27340fd444d3036ae652238f16f0309a8d35bc73066c65204d68e2640ad
crates/litchi-eval/src/parser/expr.rs              d4a657a4f268485d0143109d1d7c55ac0b7d9c25e2487b28781eafe4a01d9df7
crates/litchi-eval/src/parser/literal.rs           8c25fc4c30990869f627f17b0425757f5d9100df4191c52a760881ed8b46893a
crates/litchi-eval/src/parser/reference.rs         3fbef34a6dbb2d768663069af84886580c35e2261ef0c875ec09535d132a34ab
crates/litchi-eval/src/engine/bin_op.rs            4660fc6a3fa26bff976145b1249bf9deea1539c42eda357205ba210622cce4c3
crates/litchi-core/src/sheet/types.rs              63be6924129a3fee9bd6a6fa8701e2ed7303726a54060a897b5d7468e466db58
```

The review is a design disposition and records blockers for a future
implementation; it does not alter production sources or claim evaluation
tests.
