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
