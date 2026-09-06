# Source review

Only two standalone-tool modules change. PptxRangeSourceConfig gains an optional
NonZeroU64 transfer rate; its existing constructor and with_limits retain None,
so existing callers keep their configured behavior. Two atomics record requested
transfer sleep and paced successful calls, with the same overflow invalidation
policy as existing read counters. Snapshot/delta serialization exposes the new
values, and checked_delta refuses counter regressions. Empty/EOF/zero-cap/error
paths leave transfer counts at zero. u128 intermediates safely contain u64 byte
counts multiplied by 1e9; conversion rejects unrepresentable u64 nanoseconds.

The provider-lifecycle CLI validates an optional rate before corpus construction,
allows it only with a range provider, forwards the actual NonZero rate into the
adapter and emits the setting. Every available source/destination phase includes
cumulative and delta pacing counters; unavailable owners retain null values.
No timer boundary, budget, cache, planning, publication or output gate changes.
Exact lifecycle equivalence includes the paced source in the existing test.

Transfer sleep happens after successful underlying reads and in addition to the
existing pre-delegation fixed delay. It is per-call pacing without a shared queue
or connection model. Those limitations are explicit, and old strict report
verifiers reject the additive fields until they use the retained extended oracle.
Neither the adapter nor this experiment establishes actual network throughput.

Accepted ADR ownership and explicit-I/O contracts remain intact
(0002/0003/0005/0006/0008/0023/0024). All simulation sleeps live in standalone
tooling around caller-supplied ReadAt, with no production dependency, unsafe,
executor, global scheduling or ambient networking change. The accepted tree is
c950b6c8be822561b498d7bbe87c460873dcbf49, unchanged from its prior complete read.
