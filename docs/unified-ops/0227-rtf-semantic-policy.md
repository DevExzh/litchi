# 0227 — Checked primary RTF semantic policy

`SemanticLimits` and its consuming checked builder replace the unchecked
`edit::Limits` constructor. The profile currently covers staged operation
count. The production default remains 256; zero explicitly permits only an
empty edit. Values above `isize::MAX` return typed
`edit::Error::InvalidLimit { kind, value, maximum }` during construction.

`EditOptions` carries the checked profile into `Package::edit_with_options`;
`Edit::options()` exposes the admitted per-call policy. The old
`edit_with_limits` method and `Limits` type are removed. Runtime operation
refusals use `Error::Limit { kind: SemanticLimitKind::Operations, observed,
maximum }`, replacing `OperationLimit` in scalar, batch and picture paths.

Durable replay constructs a checked profile from the already admitted patch
operation count, preserving support for patches above the ordinary default.
Composition derives its profile from a materialized operation slice: its
non-zero-sized elements prove the count fits the addressable ceiling. These
paths do not silently revert to the ordinary edit default.

## Verification

The native suite passes 1,446 tests/doctests, including four new checked-policy
regressions and the public profile example. Coverage includes exact default,
zero and maximum construction, typed invalid input, option retention, source
immutability and reuse of an edit after operation-limit refusal. Existing
paragraph, picture, corpus, composition and durable replay suites pass with
the new typed contract.

Consumer migration is committed as `4133c8b6a2`. It updates all live native
consumer tests and the benchmark harness; historical captured source evidence
is unchanged. Final qualification passes:

- `cargo test -p litchi-ops --all-features`: 1,620 tests/doctests.
- Warning-denied all-target Clippy and rustdoc for RTF and the facade.
- RTF rustdoc JSON glossary profile: 52 banned-name checks.
- Benchmark library compilation with its copied lock.
- Crate boundaries: 66 packages; OfficeCLI capability delta unchanged.
- Glossary and OfficeCLI Python checker suites: 36 tests.

Cargo uses external copied locks (cached cc 1.6.0 for root validation);
repository lockfiles remain unchanged. The owned 0227 worktree and branch are
removed, and copied locks, build artifacts, wrappers and logs are removed after
qualification. No binary or large evidence is committed.

## Remaining work

This profile governs operation count. Existing replacement-payload admission
still uses the retained source-byte profile; a separate semantic payload
budget, admission before copying all incoming values, and focused append and
redaction vocabulary remain open. Picture-specific intrinsic caps retain
their prior behavior. Unified RTF CRUD, other tier-3 formats, Python/CLI and
full performance/corpus acceptance remain part of the active goal. This wave
makes no runtime performance acceptance claim.
