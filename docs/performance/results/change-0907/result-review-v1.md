# 0907 PPTX multi-slide result review

This is an independent offline review of the retained 0907 capture. I did not build, run workloads, run profilers, or rerun replay. The packet passes its finite evidence checks with scope caveats.

The capture is a single release binary bound to source commit `3d3650ac03be2cd8ba923a5a6e940e84b14f5eab` and source archive `69007b852482752e96bd50208891bf3dedfc4d83df7f81a3b10f42e98e64545a`. It used CPU 12, one worker, 30 warmups and 200 retained samples per leg in sequential ABBA order: eager A, source-backed A, source-backed B, eager B. All four child receipts exited zero with empty stdout and stderr, and their intervals did not overlap.

| Leg | p50 (ms) | p95 (ms) | p99 (ms) | max (ms) |
| --- | ---: | ---: | ---: | ---: |
| eager A | 263.289679 | 264.759446 | 278.711065 | 313.334012 |
| source-backed A | 7.712175 | 8.022402 | 8.174263 | 8.323143 |
| source-backed B | 7.927396 | 9.035897 | 10.655026 | 15.529992 |
| eager B | 263.581606 | 265.864489 | 266.926522 | 359.348308 |

Pooling the two legs per path gives eager p50 263.429593 ms, p95 265.490578 ms and p99 270.110281 ms from 400 samples. Source-backed gives p50 7.841286 ms, p95 8.550684 ms and p99 9.838051 ms from 400 samples. The descriptive pooled eager/source p50 ratio is 33.5952028; p95 is 31.0490457 and p99 is 27.4556699. These ratios are descriptive and have no confidence interval or release speedup interpretation.

All legs use the same corpus archive (`61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757`, 17,017,139 bytes, 445 ZIP members, 229 OPC entries) and produce the same accepted output (`23d6d7b8dd433ff453307ee56485efadc34d6b18b455a6858d9f565c4b1b6cd9`, 17,017,145 bytes). Eager reports 229 ordinary Part materializations, 741 ordinary payload read calls and 1,864 total logical reads per sample. Source-backed reports 9 ordinary materializations, 486 ordinary payload read calls and 3,372 total logical reads per sample. Read bytes are 16,907,750 / 17,030,811 for eager ordinary/all reads and 16,909,038 / 17,229,633 for source-backed ordinary/all reads. Eager sink writes are 260 calls; source-backed writes are 2,053 calls; every leg has largest write 65,536 and zero writes above that limit. These vectors are constant across the 200 samples in each leg.

The timer surrounds `publish_pptx_multi_slide_batch_edit`, including open, staging, commit, publication and the mode-specific guards. Source-backed forward-then-inverse succeeds before publication; after publication inverse succeeds and the original forward patch is expected to fail stale. Eager inverse then forward succeeds. The guards are part of the measured workflow and are not subtracted. The source-backed materialization counter is sampled after commit, before the pre-publication replay, and still pre-publication; the eager counter is the fully materialized OPC Part count. Neither is a whole-workload allocation or validation count.

`analysis.json`, analyzer mutation checks, and pre-cleanup replay all pass. The analyzer rejected malformed samples, order, percentile, materialization, read, sink and binary-identity mutations. The initial native build failure was before workloads because two compile-time fixture includes were missing; the retained native-build-v2 receipt passed with unchanged inputs.

The host was shared. CPU affinity was 12, but no exclusive host claim is established and recorded start load averages were roughly 19–21. Counters are logical in-memory reads, not physical I/O or allocator measurements. The reviewer did not reopen the PPTX or independently census raw ZIP headers, comments and member ordering, so preservation claims stay semantic/OPC scoped. SourceSummary vectors are measured-iteration order while elapsed samples are sorted; positional correlation requires `sample_order`, even though current vectors are constant. No claim follows here about absolute benchmark reproducibility, release speedup, RSS, allocation, cold-cache/device I/O, producer fidelity, no-op behavior or failed-save behavior.
