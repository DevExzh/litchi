# 0360ek — XLSB replay admission and worksheet patch raw policy

- Status: source, consumer and replay tests complete on
  `uo2/w4b-xlsb-raw-policy`; native performance qualification remains open.
- Base: `73fea85b5a` (0360ej retained raw-policy slice).
- Scope: content-matched structural replay, worksheet-only durable patch
  decoding, the affected consumers/tests, and this correction record.

## Replay admission

The earlier 0360ej implementation rejected a durable structural patch when
the profile used to decode its wire image was wider than the recipient's
retained profile. That rejected patches whose actual before/after resources
fit the recipient and therefore removed a supported content-matched replay
case.

The profile-dominance check is deleted. A changed patch still verifies its
exact before image, then builds the retained after image through the normal
owned OPC path and projects that candidate through the recipient workbook.
That projection re-admits BIFF12 records and strings under the recipient's
retained raw policy. This keeps the package publication bytes and source
provenance byte-exact while making the actual recipient admission the
decision point. The inverse path follows the same check.

Tests cover both sides:

- a patch decoded with a wider raw profile replays successfully on a narrower
  owner when the after image fits; and
- an after image with an eight-unit sheet name is rejected by a seven-unit
  recipient with the typed `raw::Error::StringLimit` observation.

## Worksheet-only durable patches

`cell_values::worksheet::Patch::from_bytes` now requires the caller's
`raw::Limits`. The wire envelope still carries its producer profile, but that
profile is intersected with the caller profile before either worksheet image
is admitted; wire metadata cannot authorize a wider read. The caller policy
is therefore required even for a detached decoder, and the effective profile
is retained by the decoded patch for subsequent validation and encoding.

The decoder test proves an exact four-unit image succeeds under a four-unit
caller policy and reports a typed `StringLimit` at three units. Existing
structural patch decode APIs already require the caller raw profile and keep
that profile through semantic operation validation.

## Audit disposition

The 0360e raw entry points remain covered by the 0360ea–0360ei slices. Public
standalone BIFF12 parsers and serializers take an explicit `raw::Limits`;
writer functions that stream through `raw::Encoder` use the encoder's already
admitted profile. `Package::from_bytes` remains the canonical owner ingress;
its policy is selected by `OpenOptions`, and detached test/builders choose
`raw::Limits::DEFAULT` explicitly. No `*_with_limits` compatibility twins or
new hidden defaults were added here. Internal `package::pivot_tables` helpers
remain crate-private and are not public raw entry points.

## Validation

All commands used the capped environment at
`(deleted-coord-artifact)`, with
`CARGO_BUILD_JOBS=1`, `CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/target-w4b-xlsb-replay`,
and a worktree-local `TMPDIR` because the shared `/tmp` inode budget was
exhausted:

- focused replay and durable-decoder tests passed;
- `cargo test --offline --locked -p litchi-xlsb --tests` passed: 759 unit
  tests and every integration target;
- package formatting and `git diff --check` passed.

The exact full-test output remains outside the checkout. No native timing,
profiler, binary or large evidence file is committed. Native qualification
and the workspace gate remain coordinator-owned.
