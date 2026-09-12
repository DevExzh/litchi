# Ordered server-format list gate

The retained final isolated run passed 47 all-feature integration tests and
12 focused unit tests, library and integration Clippy with warnings denied,
rustdoc with warnings denied, library checking, and diff checking.

`verification.json` records the baseline commit, exact source hashes, commands,
environment, lockfile hash, and raw log hashes. The accompanying `Cargo.lock`
is the operational lock used by this run. The tested files matched the shared
tree byte-for-byte when this evidence was retained. Only the final `fix2`
logs are included; earlier intermediate failures are not final-gate evidence.

Coverage includes structural insert/remove/move/reorder, required cardinality,
known reference remapping, opaque and MCE reference refusal, unchanged raw
index spelling, source-preserving inverse, and caller output limits. These
results do not establish full PivotTable lifecycle, native Excel compatibility,
refresh semantics, or runtime performance.

To reproduce, use the feature commit containing this directory, copy the
retained lockfile to the checkout root, and run the commands in the manifest
with a separate Cargo target directory. Historical absolute paths in raw logs
identify the original isolated run and need not be reused for a fresh test run.
