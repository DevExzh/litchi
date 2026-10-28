# 0393 — Pages/Keynote root-comment clear safety proof

This proof-only slice adds focused parity controls for root-only table-comment
clears on the frozen `295b421bcd3d9ec0b0b9208e4039de22623685bc` candidate.
It does not change a runtime writer, ledger row, native queue entry, or native
evidence disposition.

The tests build a root-only comment through the existing host editor over the
tracked Pages and Keynote table fixtures. The baseline clear must publish, read
back no root comment, and apply its inverse to restore the exact source bytes.
Separate negative controls append either an opaque sibling message to the
selected type-3056 object or an unknown payload field. Each control requires an
exact `UnsupportedDependency` refusal, unchanged serialized mutated bytes, and
the selected semantic root still equal after reopening. A malformed metadata
control is intentionally left to the runtime repair owner and is not conflated
with these well-formed payload controls.

## Frozen-candidate result

The baseline controls pass on `295b`. The four mutation controls are
intentionally red on that candidate: both Pages and Keynote currently publish
the corrupted clear candidate instead of refusing it. This is the preserved
regression evidence for the coordinator's W1-G repair; the proof must not be
weakened to accept that behavior.

Test module: `crates/litchi-iwa/tests/parity/comment_clear_safety_cases.rs`.
The clean evidence commit is `1a63ac0919d19d3107cc9addec533ae157ea2729`.
The scoped command was:

```text
CARGO_BUILD_JOBS=2 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH cargo test --offline -p litchi-iwa --test parity comment_clear_safety_cases -- --nocapture --test-threads=1
```

Result: 2 baseline tests passed and 4 mutation tests failed because the
candidate published a lossy clear before reaching the typed-refusal match. The
proof assertions are therefore red for the intended runtime defect. The raw
stdout/stderr receipt is outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/resume-20261028/comments-clear-proof-295b-red-raw-1a63.log`
(19,599 bytes, SHA-256
`30517c22262999d0fab50b4ff97ead59373bc95662d17bd435d49f8ab9b1c1a7`). It
records source `1a63ac0919d19d3107cc9addec533ae157ea2729`, clean status before
and after, exit `101`, and all four `unexpectedly published a lossy candidate`
failures. No performance or native acceptance claim is made here.
