# Approved paired receipts

The approved matched capture is retained in [paired-baseline-e5c18ca](paired-baseline-e5c18ca/),
[paired-candidate-251d361f](paired-candidate-251d361f/), and the deterministic
[matched comparison](matched-report.md). Both arms were captured from clean
detached control heads with the pinned Rust/Cargo 1.95.0 toolchain, then
verified from their retained raw receipts. The comparison reports only
candidate-minus-baseline deltas for the bounded lanes.

The per-arm READMEs document source/control pins, commands, binary digests,
and cleanup. Their executables and temporary checkouts were removed after
verification; the binary digest files and raw JSON, allocation, timing, RSS,
stderr, manifest, and provenance receipts remain.

The initial clean baseline is retained separately in the
[clean-f1cb11936 receipts](clean-f1cb11936/). They are bound to source pin
`f1cb119361af9ea2227d27050e41915a9a92ae04`; the clean detached checkout is
also retained at [/tmp/litchi-ink-action-baseline-f1cb](/tmp/litchi-ink-action-baseline-f1cb/).
These accepted baseline receipts are distinct from the provisional mixed-worktree
receipts described below.

## Provisional mixed-worktree receipts

These raw receipts are retained for audit only. They were captured before the
runner required every local Cargo input to match committed Git blobs, from a
shared worktree whose unrelated package and untracked files were not frozen.
They are therefore provisional and must not be cited as the final profile.
The per-process JSON, allocator counters, `/usr/bin/time -v` timing/RSS
receipts, source/build manifests, host/toolchain data, exact commands, and
verification metadata remain available for audit. The approved paired
receipts above are the isolated clean committed capture; these older files are
not part of that evidence. The `smoke/` subdirectory retains the separate
bounded correctness smoke receipts.
