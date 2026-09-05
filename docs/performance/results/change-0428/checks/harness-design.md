# PPTX managed cache-retention harness

This harness records an ownership journal for the source-backed PPTX cache.
The `lifecycle` selector uses the existing plain and media-rich cross-copy
corpora and their prevalidated semantic, package, refusal, and exact-output
gates. Corpus construction, archive digests, and the expected publication bytes
are prepared before the first sample.

Each lifecycle sample records `baseline`, `opened`, `planned`, `published`,
`drop_result`, `drop_plan`, `drop_view`, `drop_caller_sources`, and `drop_sink`.
The source and destination use distinct managed `ExecutionContext` and
`Budget` roots. At every diagnostic interval the harness calls
`SourceCacheDiagnostics::checked_counter_delta`, validates the cache gauges
against the surviving caller budget, and records the six budget dimensions
(memory, input bytes, output bytes, work, objects, and depth). Positional
`InstrumentedSource` read calls and bytes are scalar counters; they do not
attribute physical storage or decompression.

Cache diagnostics are content-free and fail closed. A consumed destination
editor and a dropped source view are represented by an explicit unavailable
state with null diagnostic fields. The final caller-source drop also makes the
source I/O counters unavailable. A cache point never turns an unavailable
owner into a zero-valued snapshot. Linux process `VmRSS` and `VmHWM` are
optional process-wide probes and carry their own scope; they are not cache or
budget measurements.

The output is opened with `create_new` and includes binary and corpus identity.
The target is untimed and makes no speed, allocator, RSS, or cache-efficiency
claim. The `exact-admission` and `one-under` media-image selectors calibrate the
managed memory floor while only the source view is pinned, then verify exact
`root_floor + leaf_bytes` admission versus a one-byte-under typed Memory
refusal. The refusal records zero source payload reads and reports the
calibration and post-refusal gauges separately; it does not infer a pure cache
wait or decompression cost.

`pinned-eviction` holds two image handles against a two-entry cache, checks a
clean eviction and normal pinned bypass, drops the handles, and verifies a
cold reload against the precomputed image oracle. `oversized-bypass` sets the
cache byte ceiling to `leaf_bytes - 1`, verifies the payload is returned and
charged to the caller while uncached, then performs a second cold read after
the first handle is dropped. `repeated-publication` reuses one source view and
caller source ownership for three exact publications, checks each output
identity, and performs a fourth publication probe that must fail with a typed
OutputBytes refusal before any sink bytes are accepted. Each source and
destination role has its own serialized configured-limit object; phase rows
retain their role-specific budget and cache availability states.
