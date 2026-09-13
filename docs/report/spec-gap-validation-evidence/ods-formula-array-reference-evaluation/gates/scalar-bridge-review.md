# Scalar bridge review

The scalar bridge is **scoped-pass** for the guards and operand profiles reviewed
here. Full value-VM integration remains **pending** and is an acceptance
blocker for the array/reference evaluator; this report does not claim that the
public value evaluator is complete.

The reviewed source is [`scalar.rs`](../../../../../crates/litchi-ods/src/codec/formula/evaluation/value/scalar.rs),
SHA-256 `4e133e7ccab7e3c4632d955a2479f2660bff30da0a494cde565b812ba4940b07`.
The final isolated test receipt reports six passing tests and status 0
([test receipt](scalar-bridge-test.json), [test log](scalar-bridge-test.log));
warning-denied Clippy also passed ([Clippy receipt](scalar-bridge-clippy.json),
[Clippy log](scalar-bridge-clippy.log)). The test context was a temporary
module wrapper ([context](scalar-bridge-test-context.txt)), so these receipts
exercise the bridge in isolation. Superseded initial compiler diagnostics were
removed after the final tests and Clippy passed.

## Confirmed bridge properties

* `eager` charges and checks execution before inspecting the stacks, the
  function name, or the argument iterator. A pre-cancelled context therefore
  cannot make a zero-argument constant or an unsupported/lazy dispatch succeed.
  The test covers the constant-cancellation case; the implementation's
  `charge_work(1)` is the common admission check.
* The bridge rejects either a nonempty scalar value stack or a nonempty frame
  stack. It leaves a caller's pending frame untouched when refusing the call,
  rather than silently consuming caller state. It also requires a
  case-insensitive match between the supplied function name and the AST
  function node. The pending-frame, mismatched-name, and constant-cancellation
  cases are covered together in the final test suite.
* `IF`, `IFERROR`, `IFNA`, `AND`, and `OR` are refused by the eager bridge.
  Their branch/sequence policy belongs to the value VM. The iterator test
  confirms that refusal happens before an argument is pulled, so a lazy or
  sequence call cannot accidentally evaluate or consume its arguments here.
* The explicit Empty profile matches the accepted local review
  ([profile](../spec-review.md)): Empty stays distinct until the consumer's
  type is known; unary `+` preserves it; numeric, logical, and text boundaries
  coerce it to `0`, `FALSE`, and empty Text respectively. Empty equality and
  ordered comparisons, formula-error precedence, concatenation, and the first
  Text/TextOrNumber parameters of the listed conversion functions are covered
  by direct tests. These comparison results are the documented implementation
  profile, not claims that Part 4 mandates unspecified Empty comparisons.
* The eager closure clears partially pushed values on argument, kernel, arity,
  and output-shape failures. The owned-text test verifies that a failed
  argument releases its reservation and that the same evaluator can be reused
  without changing the retained stack-memory baseline. A successful result is
  popped before the bridge clears its working stack, so returned owned Text is
  not dropped prematurely. Vector capacity retention belongs to the enclosing
  evaluator and was not presented as a leak test here.

## Required integration gate

The bridge source now has intended call sites in the value evaluator, but the
receipts above do not run the public `value::evaluate` path. The coder must
still wire and gate that path before treating this batch as complete. The
integration tests need to demonstrate, through a real resolver-backed value
evaluation, that:

1. unary and binary scalar operations route Empty through `Slot` without
   collapsing it to Number zero, while array element mapping preserves Empty
   and Missing distinctly;
2. scalar eager functions apply the per-function argument profile, while
   `IF`/`IFERROR`/`IFNA` remain lazy and `AND`/`OR` retain sequence omission and
   error behavior;
3. reference projection, matrix broadcasting, reference operators, and
   resolver failures do not enter the eager bridge with stale value/frame
   state;
4. cancellation, work, stack, and text-storage limits propagate through the
   public evaluator, including failures during argument production and nested
   array mapping; and
5. formula Errors remain values where the function permits them, whereas
   evaluator/resource failures remain evaluation failures and do not get
   converted into catchable formula Errors.

Until those end-to-end tests and the coder's final VM wiring receipt exist,
this review closes only the scalar bridge's local safety and profile checks.
