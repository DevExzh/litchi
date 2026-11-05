# 0212 — Canonical PPT primary transaction

`Package::edit()` now adapts the native slide-order root transaction into
`package::edit::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}`.
The family preserves the source package and publishes through the canonical
`package()` / `patch()` / `diagnostics()` triad, `into_package()` and
`into_parts() -> (Package, Patch)`. Dropping an edit is rollback.

This is an adapter over the existing native dependency-checked implementation,
not a separate presentation model. Supported native operations include slide
moves/removals/transfers, shape text and anchors, slide visibility/advance/
visual transitions, and external media properties. Native closure refusals
remain typed. The facade's existing text-edit route remains available while
unified PPT navigation and mutation dispatch are still pending.

Staging is failure-atomic: the adapter accepts a candidate native transaction
only after the operation and source-freshness checks succeed. Exact no-op
commits and patch applications retain the original package allocation.
Patches have exact byte preconditions and reversible endpoints; changed
outputs reopen under the receiving package's retained source and leniency
policies. Content-free diagnostics and redacted formatting accompany typed
errors. Native snapshot capture and publication reopens retain selected
leniency. Repaired UserEdit chains remain read-only, with the specific native
`Refusal::RepairedUserEditChain`: the append-only writer has no proven closure
for repairing that history. A cyclic-chain regression verifies tolerant reads,
the typed edit refusal, and exact source-byte preservation.

`EditOptions` defaults to a hard ceiling of 4,096 operation inputs. Its
fallible `max_operations` setter accepts stricter finite budgets; zero and
above-ceiling requests return typed `InvalidLimit`. Text batches charge each
input, successful exact no-ops consume budget, and failed calls charge nothing.
Checks precede proportional staging. The old property-set aliases under
`package::{Commit, Patch, Snapshot, Transaction}` are removed; that separate
native family remains explicitly under `package::property_set`.

## Verification

The final full all-feature native suite passes 1,385 tests/doctests, with 11
existing ignored cases. The facade suite passes 1,550 tests/doctests; all seven
focused facade PPT regressions also pass after the final native changes. Twelve
canonical regressions cover output-limit refusal and canonical transfer/text
batches, as well as exact replay/inverse and typed conflicts, projected-position composition, no-op owner
sharing, failed staging and drop rollback, retained policies, source freshness,
operation bounds, redaction, and explicit encrypted-edit refusal.

Warning-denied all-target Clippy and rustdoc pass for PPT and the facade.
All 11 minimal-feature canonical edit tests and the PPT-only facade check pass.
The PPT glossary profile passes 29 banned-name checks and verifies the six
canonical transaction-family types.

Crate boundaries pass for 66 packages and 268 internal dependencies with 11
existing debt items. All 15 OfficeCLI parity generator tests pass, and the table
is current against the base checkout's OfficeCLI source.

Reproduce with `cargo test -p litchi-ppt --all-features`,
`cargo test -p litchi-ops --all-features`, and
`cargo test -p litchi-ppt --all-features --test canonical_edit`.
Workspace checks use an external offline lock selecting cached cc 1.6.0.

## Remaining scope

This introduces the primary transaction shape, not complete PPT feature or
policy coverage. The underlying native edit admission currently refuses signed
sources, including empty edits. Thus canonical signed exact-no-op publication
is still an explicit ADR 0034 gap; changed signed writes remain refused. A
signature-specific error and exact-no-op support remain required.
Builder/new, unified PPT navigation and writes, durable facade history,
performance qualification and the all-format/Python/CLI goal remain open.
The native root currently captures a whole artifact and some operations publish
intermediate artifacts. Candidate transaction cloning also has a cost. These
are explicit measurement/optimization follow-ups; no performance acceptance
or speedup is claimed by this milestone.

The base remains an ancestor at `32ad64e3207655b4ba558fc2793d9a7396655cdf`.
Repository lockfiles remain unchanged; checks use external offline copies and
an external build target. Temporary worktrees and build artifacts are removed
after the milestone commit.
