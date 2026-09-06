# Measurement clarifications before capture

The frozen protocol's “operation peak minus entry” describes the derived
metric. The emitted `region_peak_live_bytes`, `live_bytes_before` and
`live_bytes_after` are absolute callback-order process live-byte counters.
Analysis subtracts each sample's entry from its own region peak and exit.
Raw peaks must not be compared across row sizes without retaining that entry.
Historical process highwater includes the untimed artifact/reopen; a region
peak starts at the region entry and can remain below that old highwater.

The generic operation envelope uses a `comparable_timed_operation` label.
For this batch the protocol's stricter interpretation governs: normal samples
supply descriptive writer latency, and allocator samples supply instrumented
resource observations. No latency comparison between executables is accepted.

The initial harness reader checked the expected A:D cells and row count.
Review identified that this did not reject unexpected cells outside A:D or
extra archive members. The final capture must use the strengthened oracle
checking all stored cells and exact expected package membership, with mutation
regressions. Until those checks pass, no complete-artifact acceptance is made.

The Rust/TOML/lock source manifests bind code and dependency configuration;
they do not bind every compile-time fixture in this large harness. The
streaming generator uses retained source constants and produces the named
artifacts deterministically, but the bundle is not a hermetic build kit.
Portable verification checks source-bound reports and producer oracle gates;
output archives are not exported by the existing harness.

The generic harness marks the worktree dirty when any untracked file exists.
The user-owned `docs/GOAL.md` and this evidence bundle therefore produce
`git_worktree_dirty: true`. Capture requires the tracked tree to be clean and
retains the exact before/after status in `capture-state.json`; complete code
source manifests must match the release build before and after capture. The
report's boolean alone is not a source-custody proof.
