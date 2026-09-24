# ODS scalar evaluator implementation review

Status: bounded implementation review of the direct ODS AST evaluator. This
report is separate from [`spec-review.md`](spec-review.md), which records the
normative contract and does not approve an implementation. The review is
limited to `crates/litchi-ods/src/codec/formula/evaluation.rs` and its module
registration in `formula.rs`; it does not claim complete OpenFormula
evaluation, reference resolution, array evaluation, or function-catalog
execution.

## Reviewed snapshot

The source snapshot reviewed here has these SHA-256 digests:

* `crates/litchi-ods/src/codec/formula/evaluation.rs`:
  `0975ae2c410326bc077ad3936c75be3fae65bb4e0539da4b5f24b78348e5b44e`.
* `crates/litchi-ods/src/codec/formula.rs`:
  `4e88b5ab3dc6354b1b4e12058aabd259ee87933768a7d00d99d9824a8092539d`.

The parent/tester reported the final gates passing: 879 tests in 52 targets,
two doctests, clippy, documentation, and formatting. This review did not run
Cargo commands. If either source file changes, these digests and line
references must be refreshed before treating this disposition as final.

## Findings requiring disposition

### B1: supported `TRUE`/`FALSE` calls now preserve eager child evaluation

The earlier snapshot checked the child count and directly pushed `#VALUE!` for
`TRUE` or `FALSE` with any child. The reviewed snapshot's `visit_function`
(lines 689–722) now stages an apply frame and visits every child in source
order before applying the arity error. OpenFormula §3.2.3 requires this for
functions without a specified lazy exception, and the zero-parameter `TRUE`
and `FALSE` definitions do not provide one.

This preserves child errors, capability refusals, and work limits at the
boundary. The apply frame pops values in reverse stack order while replacing
the remembered error, which retains the first source-ordered Error. Unknown
functions still fail closed before visiting children, as permitted by this
profile's explicit unsupported-function boundary. No blocker remains here.

### B2: string scans and copies now have bounded accounting

`parse_string` (currently lines 829–962) now has a checked, charged
`scan_string_body` pass for NUL and doubled-quote validation, and
`copy_string_segment` chunks and charges the decoded output copy. The second
`copy_decoded_string` discovery scan is also charged and uses
`cursor >= next_check`; advancing by two bytes cannot skip the next 4096-byte
checkpoint indefinitely. Long UTF-8 segments are copied through the checked
chunk helper.

The resulting quoted-literal work is explicit: AST visit work, one encoded
body scan, one discovery scan, and decoded output bytes/quote writes. The
formula input is bounded, and malformed quotes or NUL are rejected before
output reservation/allocation; valid UTF-8, doubled quotes, LF, and tab
semantics remain intact. No blocker remains in this path.

The same `>= next_check` form should remain in every byte loop. A modulo test
is insufficient because escaped quotes advance the cursor by two bytes.

### B3: whole-slice numeric parsing has a cancellation/accounting boundary

The numeric path retains an explicit profile boundary: `parse_number` and
text-to-number conversion charge once and then call `fast_float2::parse` over
the complete slice (currently lines 819–826 and 1161–1177). A very long
admitted numeric lexeme can therefore run without an interior cancellation
check; `fast_float2` also revisits long leading-zero mantissas, so one byte
charge is not a proof of one-pass work. The current staged foundation accepts
this as an indivisible bounded operation under the formula input cap; it does
not claim prompt cancellation or a one-pass work bound inside the dependency.
A stricter execution guarantee would require a checked lexical pass or an
explicit worst-case parser-work charge before admission.

## Invariants independently checked

The following parts of this snapshot did not present an additional blocker:

* Formula `Error` values remain `ScalarValue::Error` and propagate through
  operators; conversion failures become formula errors rather than Rust
  failures. Prefix `+` returns its operand unchanged, including its text/error
  identity.
* The evaluator uses explicit frame and value stacks. The stack capacities are
  checked before growth, vector growth is fallible, and failure drops all
  temporary reservations. No recursive evaluator walk or workbook/cache/source
  mutation is involved.
* The child storage budget accounts for live stack and text reservations and
  participates in ancestor-budget accounting. Owned text moves its reservation
  into `EvaluatedScalar`, so the result remains accounted for until dropped.
* Concatenation checks decoded result length before allocation, uses fallible
  growth, and geometrically grows an owned left buffer. Its stored-cell path is
  therefore amortized for flat concatenation; this review makes no whole-parser
  complexity claim.
* Numeric non-finite results become `ScalarError::Number`; division by zero is
  a formula error; the documented `0^0 = 1` choice is allowed by §6.16.46.
  Text equality/ordering uses the documented case-sensitive UTF-8 scalar
  ordering profile, and mixed ordered types are returned as a typed value
  error.
* References, arrays, names, labels, missing arguments, and unimplemented
  functions fail as typed capability refusals. This is the intended staged
  boundary; the two functions advertised as supported now visit their
  arguments before reporting arity errors.

The report does not claim that the current profile implements the remaining
host-dependent conversions, references, arrays, lazy functions, or the full
OpenFormula function catalog.
