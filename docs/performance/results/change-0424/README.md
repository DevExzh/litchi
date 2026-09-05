# 0424: reuse validated staged PPTX payloads

Publication previously cloned the selected image/chart payloads a second time
when preparing an existing source-backed cross-copy plan. The candidate shares
the plan's independently staged immutable bytes after exact metadata and byte
comparison. Full source validation, final plan equality, candidate reread,
cancellation checks and conservative memory reservations remain.

The production revision is `d18bf7db4`; the control is
`6ca9962c7818e173538a40ed45f3a3c32cc1aa6a`. The [design](design.md) includes the
ADR and ownership review. Both [allocation-profile](protocol.json) and
[matched-measurement](measurement-protocol.json) protocols were frozen in
`5b176dc9a` before implementation. The benchmark harness is unchanged.

The control uses the exact retained normal and allocator binaries from the
0423 clean build. `build-origin.json` retains that build receipt; the control
receipts explicitly identify artifact reuse. The candidate uses a clean Rust
1.98.1 release build with matching flags. Reports bind source, binary, corpus,
output and verifier identities. The formal matrix contains 16 fresh processes
and 1,040 retained observations: 100 samples / 10 warmups in the normal lane,
30 / 3 in the allocator lane, two corpora and two repeats per role. Each group
runs control R1, candidate R1, candidate R2, control R2 on CPU 2 with one worker.
All task CPU workloads are serialized; other activity on the shared KVM host
is uncontrolled.

See the [complete result table](matched/result-table.md),
[machine-readable summary](matched/summary.json), [resource decision](resource-review.md),
[stack attribution](profile-notes.md), [validation](validation.md),
[source review](checks/source-review.md) and [remaining work](next-work.md).
All 16 report replays, four trace replays and 104 mutation probes pass in a
standalone export after removal of the original worktrees and copied binaries.
Normal timing is diagnostic, with no release latency claim. Allocator request
volume, absolute region live-byte maximum, live endpoints, logical read counts
and whole-process RSS have distinct scopes. Neither callback accounting nor
whole-command Heaptrack totals measure physical copies or post-drop retention.

Replay with Python 3, Heaptrack and zstd available:

```sh
python3 -B docs/performance/results/change-0424/portable-replay.py
```

Portable replay exports this bundle, its four hash-pinned validators and the
hash-bound historical allocation parser into a temporary directory. It checks
the 16-report deterministic summary, eight R1 report mutation suites, both
corpora's control and candidate allocation traces, and rejection of a modified
pinned validator. It needs neither original worktree nor copied binaries.
The replay receipt is updated, so use a copy to preserve the final inventory.

Raw report, catalog, journal and trace bytes remain unchanged. Historical
failures are explicitly separated. Lossless compressed logs are bound through
`compression.json`; `SHA256SUMS` inventories the bundle. Build and capture
scripts retain their exact command contracts and refuse existing outputs;
recapture into a fresh directory with clean source checkouts and the matching
builds. This batch does not complete the broader non-iWork goal.
