# 0428 next work

0428 is a measurement enabler. The additive PPTX fallible cache-diagnostic
forwarders and their focused source/no-payload-read checks do not report a
cache, budget, RSS, retention, allocator, latency, or optimization result.
The overall non-iWork goal remains open. The [protocol](protocol.json),
[cache design](checks/cache-design.md), and [harness design](checks/harness-design.md)
are planning artifacts until the formal capture and independent replay pass.

## Highest-value immediate measurement

Run and verify the complete frozen 0428 managed-lifecycle matrix. This is the
smallest measurement that can turn the new seam into useful evidence because
it binds cache gauges and caller budgets to the real source-backed PPTX
cross-copy ownership boundary. It should retain:

- 16 fresh processes, CPU 2, one worker, three warmups, and 30 samples per
  process, for 480 retained samples;
- plain and media-rich lifecycle rows;
- media-rich exact-admission, one-byte-under, pinned-eviction, and
  oversized-bypass rows; and
- plain and media-rich three-publication rows under the cumulative output
  ceiling.

The capture must use separate source and destination `ExecutionContext` and
`Budget` roots. It must preserve the consuming destination editor's final
pre-publication snapshot and represent its post-publication state as
unavailable, rather than fabricating zero. Source snapshots, both caller
budgets, positional source counters, cache monotonic counters, and point-in-
time gauges must remain separately scoped through the named drop phases.

Every diagnostic read must use the new `try_cache_diagnostics` path. Every
same-owner counter interval must pass
`SourceCacheDiagnostics::checked_counter_delta`; an error, missing owner,
reordered phase, or counter regression invalidates that row. Exact and
one-under rows must retain their metadata calibration, selected payload
identity, typed result, source-read delta, zero-output refusal check, and
cache/budget gauges. Pinning, eviction, and oversized bypass must be
distinguished by their actual handle ownership and cache counters. Repeated
publication must show the expected releasable `Memory`/`Objects` behavior
without treating cumulative `InputBytes`, `Work`, or `OutputBytes` as
releasable memory.

The independent verifier should retain source revision, binary, corpus and
output identities, all cache and budget limits, raw phase journals, and
failure rows. It should report process RSS/VmHWM only with its process-wide
scope. It must not turn these observations into a timing comparison, physical
copy count, leak proof, cache-efficiency result, or causal optimization claim.

## The next gap after this capture

If the frozen matrix replays cleanly, the highest-value next measurement is one
matched native-producer and cold/range-source extension of the same lifecycle
protocol, beginning with the media-rich source-backed case. The current 0428
corpora are deterministic synthetic in-memory PPTX inputs; they can establish
the named cache and budget transitions but cannot establish behavior for
producer-specific package layouts, filesystem cold state, or caller-supplied
high-latency ranges. Reuse the same phase names, independent budgets,
fail-closed diagnostics, output/preservation gates, and explicit unavailable
owner states so the extension remains comparable. Keep native, cold, and range
effects separately labelled rather than combining them into a single
optimization result.

Broader CRUD coverage, non-PPTX format lifecycles, bounded streaming/append,
parallel scaling and hardware-counter attribution remain global requirements.
They are not closed by the 0428 cache/budget enabler or by a successful
descriptive capture.
