# 0193 — DOC strict-owner source admission

## Change

Follow-up to record 0192: the DOC body editor must check retained physical
source limits after parsing CFB directory metadata and before the strict OLE
owner captures payloads. The shared owner supplies a statically dispatched,
caller-owned admission callback with an immutable parsed container and typed
errors. Admission reuses the same parse that the owner returns to DOC.

Ordinary and profiled body opening, edit construction, auxiliary candidate
reopening, and changed candidate publication retain the selected profile.
Outer input checks precede parsing; declared entry and aggregate checks precede
payload capture. Reader validation still runs afterward. Parse diagnostics
report container parsing separately from a format-policy refusal.

## Verification

The combined native DOC/shared OLE/XLS/PPT and legacy facade suite passed
5,276 tests and doctests, with 28 ignored. After adding the tightened-profile
input gate, the DOC/facade suite passed 1,913 tests and doctests, with 14
ignored. Entry and aggregate error-precedence tests prove that source-policy
refusal precedes the strict owner's capture gate. Profiled refusal retains one
successful CFB parse pair and closes the strict-validation phase with an error.
An auxiliary candidate with invalid CFB bytes above the selected input cap
returns the input-limit error before parsing. Facade oversized-edit regressions
preserve typed entry/aggregate dimensions and the original bytes/text.

Warning-denied all-target Clippy and rustdoc pass; the common admission
example also executes successfully as a doctest. Both DOC glossary profiles
pass 34 banned-name checks and one transaction-family check. The crate-boundary
audit passes for 66 packages and 268 declarations, with 11 existing debt items.
Scoped formatting and diff checks pass. Cargo uses an external copied lockfile
with cached cc 1.6.0; repository lockfiles remain unchanged. No iWork crate is
modified. The rewritten base remains an ancestor of this branch.

## Remaining scope

Physical limits bound declared input and logical stream sizes. They do not
bound all live copies or total process memory. Already-parsed raw adoption
cannot retroactively constrain allocations made by its caller; low-level CFB
directory and allocation-table limits remain separate dimensions. Standalone
focused editors with their own explicit limits retain those contracts.

This closes the strict body-owner ordering follow-up from record 0192, not the
full unified-operations goal. Broader DOC CRUD, remaining focused/raw vocabulary,
XLS/PPT write loops, tier 3, Python/CLI and full performance acceptance remain
open. No performance improvement or allocation-count claim is made here.
