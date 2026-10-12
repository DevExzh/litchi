# 0226 — RTF atomic publication

Native `Package::save`, `save_with_durability` and `save_with_options` publish
through the new format-neutral `litchi_core::atomic` owner. Ordinary save uses
full durability and replacement. `save::Options` supplies explicit durability
and atomic no-clobber policy. All three methods return the serialized byte
count and leave the immutable package and its opening profile unchanged.

The shared publisher stages serialization in a sibling file, flushes it,
preserves existing permissions, performs requested synchronization, verifies
its temporary identity and destination type, then publishes. No-clobber uses
an atomic hard link. Pre-publication errors preserve the destination; failures
after publication retain a typed `Committed` error. Cleanup authenticates the
temporary pathname before unlinking it. Portable pathname checks require a
trusted parent directory; they do not make a hostile directory race-free.

Unified RTF supports save to an explicit path, save to a retained opening
origin, and stream output. Saving a copy does not change the origin, and
created/byte-opened handles require an explicit destination. Dispatch adopts
the native publication policy and preserves its typed error chain. Capability
reporting now advertises publication separately from pending CRUD navigation.

The CFB, OPC and flat ODT publishers are unchanged by this wave. The new core
publisher provides a format-neutral dependency for RTF and future consumers;
this record does not claim consolidation of the earlier publishers.

## Verification

The core suite passes 307 tests/doctests, including injected synchronization,
verification, substitution and committed-cleanup failures. The native RTF
suite passes 1,441 tests/doctests. Facade qualification also exercises all three
durability levels, exact navigation-fixture bytes, native edited publication,
stream failures, explicit and retained origins, no-clobber, symlink refusal,
permission preservation and the typed committed-error chain.

The all-feature facade suite passes 1,619 tests/doctests. The final RTF-only
suite passes 15 checks (creation, ingress and seven publication regressions),
including the navigation/edit regression added after the full facade run.
All-target warning-denied Clippy passes for core, RTF and the facade.
Crate boundaries pass for 66 packages; the OfficeCLI delta remains unchanged;
all 36 Python checker tests pass.

The publisher foundation and review fixes are committed as `556a4ffd44`,
`7e96615ead` and `925003fa0e`. Warning-denied rustdoc passes for all three
crates. The RTF glossary profile passes 49 banned-name checks; facade
compilation with no features also passes.

Tests run on Linux. Windows code and its readonly-cleanup regression are
present but were not executed here; only the Linux Rust target is installed.
Unsupported non-Unix/non-Windows platforms refuse publication before staging.

Cargo uses an external copied lock with cached cc 1.6.0; repository lockfiles
remain unchanged. Temporary worktree, copied lock, logs and build artifacts
are removed after qualification. No binary or large evidence is committed.

## Remaining work

RTF semantic profiles, focused transaction vocabulary, unified CRUD and broader
authoring remain required. The all-format goal, Python/CLI completion and full
performance/corpus acceptance remain active. No runtime performance acceptance
is claimed by this wave.
