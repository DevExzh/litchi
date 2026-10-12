# 0240 — Typed RTF body-paragraph selectors

## Contract

The remaining primary RTF body-paragraph edit selectors now use
`litchi_core::Position`. This completes the body-edit selector migration started
by typed insertion, swap, alignment batches and layout batches. The scalar
replacement, insertion-after, removal, move, split, merge and alignment methods
accept typed source-relative positions; bulk removal accepts a slice of typed
positions. `ParagraphTextReplacement` stores and returns `Position`.
`TransferPlan::plain_paragraph` uses typed source and destination positions.

The duplicate `split_paragraph_at` entry point is removed; callers use
`split_paragraph`. Its UTF-8 byte offset remains `usize`, as do text-span byte
coordinates, counts and native diagnostic fields. The immutable source still
owns range validation. No compatibility aliases or implicit integer entry
points are added. The glossary gate bans the removed alias.

Native operations and durable patch wire coordinates retain their existing
representation and validation. The facade retains typed positions from target resolution through native
staging, without an additional conversion vector. Benchmark boundaries convert
raw ordinals explicitly. Native tests, examples and tools use the
same typed vocabulary.

## Verification

The native suite passes 1,495 tests/doctests, the full-feature facade passes
1,697 and the RTF-only facade passes 562. Rustdoc examples cover typed body selection with a UTF-8 byte
boundary and reject raw integers for replacement descriptors and removal
batches. The existing native lifecycle, source-policy, budget, no-op, history,
durable replay and source-preservation tests use the new API directly.

A rustdoc JSON audit verifies all twelve body selector methods, the three
paragraph update constructors/accessors and plain paragraph transfer use
`Position`. It also verifies the split offset remains `usize` and the alias is
absent. The updated glossary policy rejects the pre-migration rustdoc JSON for
`split_paragraph_at` and passes the migrated API (86 banned names, one
transaction family).

Warning-denied all-target Clippy passes for the native/facade all-feature build
and RTF-only facade. Warning-denied rustdoc, no-feature compilation and scoped
format/diff checks pass. The three repository-tooling test modules pass 1,166
checks; crate boundaries pass for 66 packages and 268 internal dependencies,
and the OfficeCLI delta is unchanged. All targets of the main performance
harness compile; the standalone RTF removal harness passes warning-denied
Clippy. All nine RTF-filtered performance-harness tests pass, including the
frozen large-output hashes and paired native/unified read oracles. Two stale
native commit accessor calls in benchmark tests are migrated
from `snapshot()` to the canonical `package()`.

This is a selector-contract migration, with no new timing qualification claim.
Previous measured performance gaps remain open. Builds use an external copied
lockfile with cached `cc 1.6.0`; repository lockfiles are unchanged. Both owned
agent worktrees/branches and owned scratch/build artifacts were removed after
verification. No binary or large evidence is committed.

## Remaining work

Retained-story selectors and other native glossary debt remain outside this
body-edit migration. RTF character properties, richer mixed text/layout writes,
formatted paragraph creation, broader elements, repeated-command coalescing,
allocation/scaling qualification and earlier no-op/durable swap gaps remain.
Other delayed formats and Python bindings are still required by the full goal.
