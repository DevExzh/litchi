# Deferred scalar-cell review

Reviewed `value.rs` SHA `cf2f0dee3a44bb2bc1830526b1d4942b1a718f2ef3dae55da2b9f965625fa1b9`,
`references.rs` SHA `1778cd9baced1a6fd69ca0708fad4540e5c031cf27c8fca423c76fd08b50f04c`,
and the owned conversion guards in commit `6770a4d36`.

The independent reviewer found no confirmed correctness blocker. Only exact
single-cell scalar-demand references create the private token. Matrix mode,
reference operators, and generic arguments retain first-class areas. Endpoint
metadata and extent/admission checks still precede projection; projection keeps
the current-sheet probe and left-to-right reads. Cumulative read limits,
cancellation checks, and the outer source-version fence remain in place.

Publication and cache paths reject or exclude tokens: `finish_root`, owned
measurement/copy, demand cache, and condition cache explicitly handle them. The
fallback in borrowed inspection is unreachable under the publication invariant.
A missing current-sheet result retains the prior single-area behavior.

[Full gates](scalar-cell-candidate-01.json) pass on the reviewed source closure.
The paired performance capture separately identifies a one-cell matrix-range
regression; this source review does not close performance acceptance.

The final variant-order-only follow-up passes [all five gates](scalar-cell-candidate-02.json)
with the same 455-file closure checks. It moves the new private token after the
existing array/reference variants without changing the reviewed matching behavior.
Paired capture03 retains zero deterministic mismatches; a 5–6% one-cell matrix-range
p50 regression remains open.
