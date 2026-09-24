# Bitwise evaluator implementation review

Status: closed for the reviewed bitwise candidate. This is an independent
correctness and resource review of the five OpenFormula §6.6 scalar functions;
it does not claim reference, array, workbook, or recalculation support.

## Reviewed inputs

* Base implementation: commit
  `c32455df83ea67f840a6d4f8da8953f113b590f4`.
* Candidate source:
  `crates/litchi-ods/src/codec/formula/evaluation.rs`, SHA-256
  `a150855aa43b64753ba90ba98322b213d1dfa0bc33b0e317da0f3e7e11a959db`.
* Focused test source:
  `crates/litchi-ods/tests/ods_formula_bitwise_evaluation.rs`, SHA-256
  `a64fcb6976541001c3ce33a51e9bfa44283717db14ea61cf853dd5325a65d8b8`.
  It contains nine focused integration tests covering truth tables, shifts,
  conversion and domain limits, formula-error order, capability refusals,
  lazy handlers, resource/cancellation behavior, and nested evaluation.
* Normative artifact:
  `3rdparty/specs/OpenDocument-v1.4-os.zip`, SHA-256
  `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`;
  entry `part4-formula/OpenDocument-v1.4-os-part4-formula.html`, SHA-256
  `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.
* The normative function and profile decisions are recorded in
  [`spec-review.md`](spec-review.md). Resource and failure boundaries were
  checked against accepted ADRs 0001, 0005, 0006, and 0008.

No Cargo command, build, or production edit was performed for this review.

## Semantic disposition

`BITAND`, `BITOR`, and `BITXOR` convert both operands through the existing
Number conversion path, truncate finite values toward zero as the selected
§6.3.6 profile requires, check the unsigned inclusive `0..=2^48-1` domain,
and apply the corresponding `u64` operation. The result check remains in
place even though valid operands make these three operations fit naturally in
48 bits.

`BITLSHIFT` and `BITRSHIFT` use the same 48-bit data-operand conversion and a
signed, truncated finite shift count. Negative counts reverse direction, zero
counts preserve the operand, right shifts of magnitude at least 48 return
zero, and left shifts of zero return zero for any finite count. A nonzero left
shift with magnitude at least 48 returns `#NUM!` under the documented profile.
For smaller counts, `shift_left` checks `value <= BIT_MAX >> amount` before
the native operation and then uses `checked_shl`; this prevents high bits from
being discarded and accidentally producing a plausible low result. The
count is cast only after the `< 48` comparison, so a huge finite count cannot
truncate into a platform shift mask. All shift paths are constant time and
avoid exponentiation, loops, negation of an integer minimum, and shift panic.

The conversion guard uses the exact `2^48` floating boundary, so `MAX48` is
accepted and `2^48` is rejected. Truncation precedes the nonnegative check,
which intentionally accepts `-0.9` as signed zero and rejects `-1.9`.
Non-finite numbers and failed text conversion remain formula errors according
to the existing scalar conversion policy.

## Evaluation and error behavior

The five names use the existing eager scheduling path. Every supplied child
is visited before application, including a malformed-arity call. The
two-argument check then drains the value stack and returns `#VALUE!` when no
child formula Error exists. Reverse stack popping with repeated assignment
retains the leftmost source-order formula Error. Formula Errors therefore
remain values, while unsupported references/arrays, cancellation, resource
limits, allocation failures, and invalid evaluator state continue as typed
`EvaluationFailure` results.

`IF` still avoids an unselected bitwise branch, and `IFERROR`/`IFNA` catch only
the selected bitwise formula errors. They do not catch evaluator failures.
The bitwise application itself introduces no branch that weakens eager
evaluation or error propagation.

## Resource disposition

The implementation reuses the existing explicit `Frame::Apply` and value
stack. Child scheduling charges work before traversal and checks the caller's
execution context; each frame visit and application retains the existing
work/cancellation checks. Text conversion charges the consumed UTF-8 bytes
through the existing bounded path. The bitwise helpers move values and use
fixed-width scalars only: they allocate no heap storage, retain no storage
reservation, and perform no workbook, cache, package, filesystem, network, or
ambient-provider operation. Existing fallible frame/value stack reservations
remain the only evaluator-owned allocation path.

The nine focused tests include cancelled and zero-work/zero-storage refusals,
typed capability failures under wrong arity, high-bit left-shift overflow,
huge positive and negative shift counts, 48-bit boundaries, and nested calls
under explicit stack/work limits. The retained performance harness describes
the candidate-only bitwise lanes and baseline refusal control; its report
does not claim timings until root capture receipts are present. The retained
crate-boundary gate receipt passes.

## Verdict

No correctness, boundedness, or resource-accounting blocker was found in the
reviewed source/test vector. The high-bit pre-shift guard is required for this
verdict and is covered by the focused overflow fixture.
