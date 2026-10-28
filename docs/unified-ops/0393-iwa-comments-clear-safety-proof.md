# 0393 — iWork root-comment clear safety proof

This proof-only slice adds focused parity controls for root-only table-comment
clears on the frozen `295b421bcd3d9ec0b0b9208e4039de22623685bc` candidate.
It does not change a runtime writer, ledger row, native queue entry, or native
evidence disposition.

The tests build a root-only comment through the existing host editor over the
tracked Pages and Keynote table fixtures. The baseline clear must publish, read
back no root comment, and apply its inverse to restore the exact source bytes.
Separate Pages/Keynote negative controls append either an opaque sibling
message to the selected type-3056 object or an unknown payload field. Numbers
controls additionally exercise a surviving unknown ArchiveInfo reference to
the selected root, a selected-storage `should_merge` header mismatch, and a
nonempty malformed type-6000 candidate. Each control requires the old typed
refusal contract, unchanged serialized mutated bytes, and the selected
semantic root still equal after reopening. A malformed metadata control is
kept distinct from these proof cases.

## Frozen-candidate result

The baseline controls pass on `295b`. The eight mutation controls are
intentionally red on that candidate: Pages, Keynote, and Numbers currently
publish the corrupted clear candidate instead of refusing it. This is the
preserved regression evidence for the coordinator's W1-G repair; the proof
must not be weakened to accept that behavior.

Test module: `crates/litchi-iwa/tests/parity/comment_clear_safety_cases.rs`.
The clean evidence commit is `88237f11d82353f0ec1666267ad34ec9e39e30c3`.
The scoped command was:

```text
CARGO_BUILD_JOBS=2 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH cargo test --offline -p litchi-iwa --test parity comment_clear_safety_cases -- --nocapture --test-threads=1
```

Result on the extended proof run: 3 baseline tests passed and 8 mutation
tests failed because the candidate published a lossy clear before reaching the
typed-refusal match. The proof assertions are therefore red for the intended
runtime defect. The fresh raw stdout/stderr receipt is outside the repository
at `/home/zhuhe/code/litchi-worktrees/.coord/resume-20261028/comments-clear-proof-295b-red-raw-extended2.log`
(21,443 bytes, SHA-256
`fa948bb36469434058d03fc5403e042c54dd8d6af615891fca79188c865b23f2`). It
records source `88237f11d82353f0ec1666267ad34ec9e39e30c3`, clean status before
and after, exit `101`, and all eight `unexpectedly published a lossy candidate`
failures. No performance or native acceptance claim is made here.

## Repaired shared-engine result

The exact proof tip `d50d57d8c439d2891f81311383b95d30eeb3a663` was merged into
the repair WP at `d34aacac9a`. The repaired shared engine adds the strict
package-wide reference/header/table admission described in record 0392. The
current parity module contains 11 cases (three successful baselines and eight
typed loss/refusal controls), and all 11 pass on the repaired source:

```text
CARGO_BUILD_JOBS=2 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 \
CARGO_PROFILE_TEST_DEBUG=0 cargo test --offline -p litchi-iwa --test parity \
comment_clear_safety_cases -- --nocapture
```

The repaired run preserves source bytes and semantic roots on every refusal;
the successful baselines reopen and invert byte-exactly. This remains parity
proof only: no native E3/E4 acceptance or performance qualification is
claimed.
