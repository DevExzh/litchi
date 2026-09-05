# Evidence driver review

Source review checked the capture argument positions, report/binary/protocol
identities, repeat ordering, compression fallback, profile artifact set, and
portable replay flow. The replay wrapper now checks both positive and shapes
static native oracles and pins its own digest through the build record.

The complete Rust/TOML/lock manifest is checked before binding the release
build and before/after capture; this includes untracked source files, unlike
`git diff` alone. The check driver is pinned in the build record. CLI smoke
results remain preflight controls outside formal sample counts. Full capture
reports receive independent validation and mutation replay.

Preflight exposed two independent driver assumptions: the generic synthetic
manifest belongs to the destination, and unavailable baseline histograms are
null. Both corrections preserve the explicit source/destination identity and
unavailable-owner contracts. Rejected reports and logs remain retained.

Source review alone is not execution evidence. Final receipts, indexes and
portable replay establish which captures completed. Profile sample attribution
is descriptive CPU ancestry, with setup/observer and unresolved frames retained;
it is not an API wall-clock fraction or causal speedup.
