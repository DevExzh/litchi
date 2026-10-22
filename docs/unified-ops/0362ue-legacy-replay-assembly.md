# 0362ue — legacy replay assembly validation

The first combined run on `d3b36e8c74` failed one DOC unit assertion:
`protected_package_allows_noop_but_requires_explicit_authorization_for_metadata_edit`.
The remaining 1,228 unit tests passed and two existing tests were ignored.
The run stopped before its remaining six groups.

The raw-conflict migration had changed the assertion for the original,
matching protected artifact to expect `PatchConflict`. Production correctly
returns `ProtectionDenied` because this replay supplies no authorization.
The corrected test preserves that distinction and separately checks that the
changed artifact returns `PatchConflict`. Authorized replay must still produce
the exact committed artifact. No production policy or error precedence changed.

Direct rustfmt passed. The combined tests and affected-consumer checks will
be rerun on the committed correction. The failed run is retained externally:

- Log: `/home/zhuhe/code/litchi-worktrees/.coord/replay-integration/legacy-replay-combined-v1/check-0.log`
- Size: 229,601 bytes.
- SHA-256: `81e322c5d5a8b34fa3a6b97814dcb7f8a9eadad4a9835ba97cb42b902839d57a`.
- Command: `cargo test --offline --locked --all-features --message-format json-render-diagnostics -p litchi-doc --lib`.
- Exit: 101; duration: 41.455 seconds, including compilation.

This receipt does not claim the final workspace gate or native performance
qualification.

## Corrected DOC run and facade compilation repair

The second run held source `5fbcfc45e7c5611f75d32968154a7aa11f6508c0`
fixed. Its first three command groups passed 1,487 tests: 1,229 DOC tests,
227 OLE common unit tests, 15 OLE1 integration tests, and 16 DataSpaces
tests. Three existing tests were ignored across those suites.

The fourth group, facade error classification, stopped during compilation.
The new DOC test had imported `Position` at module scope, making three
existing qualified uses trigger the workspace's denied `unused_qualifications`
lint. Removing that import and qualifying the two new test expressions
preserves every feature combination without changing production behavior.
Direct rustfmt passes for the correction. The facade group and remaining
DOC facade, Clippy and harness groups still need to run; the host disk fell
below the coordinator's 12 GB free-space floor before a rerun could start.

The exact commands, durations, log sizes, hashes and test summaries are in
[validation-v2.json](results/0362ue/validation-v2.json). The largest target
size recorded before a command was 18,845,543,576 bytes. The runner hashed
and removed only completed debug test executables; its external receipt
retains those hashes. Earlier failure evidence remains valid and retained.

## Completed combined run

All seven groups passed on the fixed source
`c2122e26786f02058b36a95a0162c42489a3ad65`, with the pinned Rust 1.95.0
toolchain. The run covers 1,578 tests: the 1,487 DOC/OLE/DataSpaces tests
above, 78 facade error-classification tests, and 13 DOC facade tests. Three
existing tests remain ignored. Strict all-target/all-feature Clippy passed
for DOC, OLE common, crypto and ops; the locked all-target/all-feature
performance-harness check also passed.

[validation-v3.json](results/0362ue/validation-v3.json) records every command,
test summary, duration, external log path, size and SHA-256. The largest
pre-command target size was 18,982,780,844 bytes. The memory-capped runner
used one Cargo job, no incremental compilation, debug information disabled,
and a temporary directory outside the host `/tmp`. Completed debug test
executables were hashed before removal; their hashes remain in the external
runner receipt. The earlier failed runs are retained.

This result includes the consuming DOC handoff through `76391f1a71`, but
does not qualify its performance or newer worker source. It is an affected
assembly result, not the final workspace gate, corpus disposition check,
or native-performance acceptance.
