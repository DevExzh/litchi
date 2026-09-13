# Value evaluator integration checkpoint 14

This checkpoint wires the bounded value evaluator, reference operators and
borrowed worksheet resolver into the ODS crate. It is an implementation
checkpoint with **two known failing regressions**, not production acceptance.
The original scalar evaluator remains a separate path.

The resolver API now passes the evaluation caller's `ExecutionContext` into
every provider operation. Worksheet index preparation retains its own memory
reservation; later lookups charge their supplied context. Tests cover a reused
index with independent preparation and lookup policies, cancellation, retained
reservation release, and the dedicated `UnsupportedKind::CellValue` refusal for
formula caches and Date/Time/Unknown cells.

The exact isolated snapshot passed 390 library tests and 16 worksheet tests.
The value suite passed 41 of 43 tests. Its remaining failures are:

- `nested_if_scalar_branch_uses_selected_condition_shape_only`: a nested
  two-column condition with a scalar selected branch produces 1×1 instead of
  1×2. Condition shape must participate in nested result planning.
- `provider_driven_nested_handlers_keep_selected_shape_and_skip_unselected_metadata`:
  a provider condition is read three times instead of once. Planning and output
  evaluation must reuse the same demanded-coordinate decision.

These failures remain enabled. The broader planner review also requires correct
demand propagation when selected shapes grow, stable cache identity and bounded
nested evaluation; passing these two examples alone will not establish that
contract. The selected-branch 3×2 broadcast regression passes in this snapshot.

The [receipt](../performance/diagnostics/integration-checkpoint-14.json) and
[verified archive](../performance/diagnostics/integration-checkpoint-14.tar.gz)
retain commands, toolchain, environment, both test logs and SHA-256 hashes of the
452-file build source closure. The checkpoint's production files were staged
from those tested isolated bytes while the coder continued later changes in the
working tree. The loose archived logs were removed after verification.

Supplemental checks on this same snapshot subsequently passed warning-denied
all-target Clippy, warning-denied rustdoc, all four doctests (including the
compile-fail lifetime example), and formatting. Their complete logs are in the
[supplemental receipt](integration-checkpoint-14-extra-checks.json). Full runtime
gates and release performance captures remain pending for the final
implementation. The
[API integration requirements](value-api-integration.md), including structural
view equality, explicit limit contracts and budgeted owned conversion, also
remain open. No feature matrix is promoted by this checkpoint.
