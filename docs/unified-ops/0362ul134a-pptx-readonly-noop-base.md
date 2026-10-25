# 0362ul134a — PPTX readonly no-op replay base

Status: source repair committed; independent behavioral review ready.

## Diagnosis

The full-feature runtime test `replay_rechecks_readonly_output_with_the_recipient_profile`
constructed its no-op patch from `Package::new()` and then applied that patch to
a permissive package whose `/ppt/presProps.xml` had been replaced with the
large opaque fixture. Exact no-op patches are content-rooted, so the resulting
`PatchConflict` was the expected foreign-base refusal, not a replay or physical
limit failure.

## Repair

The test now creates the no-op edit from the permissive recipient itself. It
replays successfully on that recipient and on the equal-content recipient
opened with the tighter `SourceLimits` profile, asserting byte preservation in
both cases. It also retains an explicit foreign `Package::new()` replay,
asserting `PatchConflict` and unchanged bytes. The strict changed-patch
physical-limit path is unchanged.

The parallel math and presentation-metadata no-op/replay consumers were
audited; they derive and apply their no-op patches on the same source base and
needed no production or test migration.

## Verification

- Base commit: `780cf4c7eebd62784670e2ab14cc49088a4bc91c`.
- Behavioral fix commit: `52bdea122df3bc852b52ba37f95324ce93dce951`.
- This follow-up removes the unused `mut` rejected by strict warnings.
- `rustfmt --edition 2024 --check crates/litchi-pptx/tests/pptx_readonly_recommended.rs` passed.
- `git diff --check` and repository commit size/ELF guards passed.
- Cargo, build, runtime tests, native tools, rustdoc and metadata generation
  were not run in this source-only repair lane.
