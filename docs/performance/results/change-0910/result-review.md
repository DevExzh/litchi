# 0910 result review

Decision: **approved with no blocker** for the frozen descriptive report. This
review independently read the 24 retained child reports and receipts, rebuilt
the pooled summaries and paired bootstrap from the raw sample vectors, checked
the order split and process RSS, and inspected both raw ZIP census reports.
No build, benchmark workload, replay, or profiler was run for this review.

The final report is bound to SHA-256
`3c155ffd3a97a2c9f9bb99eb8ceb21d2ef1d42bfc581f921d5a91e41c14ef5e6`.
The canonical reducer is bound to
`62f8dcd7ad6f226d1e5d3527bd6afdcce5ca12775d1a826bb5b6fda105f7acf2`.

## Independent numerical review

The raw packet contains 24 successful child processes, 12 AB/BA pairs, and
480 retained measured samples. Each child records 30 warmups and 20 measured
operations. The source-bound harness verifies each warmup and measured output;
the packet therefore accounts for 1,200 guarded executions, while retaining
timing vectors for the 480 measured operations. Every retained sample passed
the output digest, output size, sink, source and phase arithmetic checks, and
all child processes exited zero.

The independently recomputed pooled p50 values and paired process-median ratios
are:

| Phase | Materialized p50 | Owning p50 | Materialized / owning paired median |
| --- | ---: | ---: | ---: |
| Total | 259.703332 ms | 32.180007 ms | 8.072147309 |
| Clone | 0.810824 ms | 2.903215 ms | 0.288418782 |
| Ingress | 6.194567 ms | 2.971650 ms | 2.103250647 |
| Opened snapshot | 12.349500 ms | 15.865198 ms | 0.777935667 |
| Edit and commit | 2.868150 ms | 2.942265 ms | 0.977555984 |
| Publish, serialize and sink | 236.840172 ms | 5.575988 ms | 42.497503985 |

The total paired bootstrap is 10,000 resamples of the 12 complete process
pairs with `random.Random(910)` and nearest-rank percentiles. Its 95% interval
is **7.960537979–8.240713911**. The publication phase interval is
**41.593129673–43.239137170**. The clone phase is noisy: its paired interval
is **0.273646303–0.914021815**, and the same `source.to_vec` operation is
exposed to allocator and cache state. The report correctly discloses the
owning clone and snapshot p50 increases; the clone p50 should not be treated
as a route mechanism or an isolated causal cost.

Order drift does not identify a material order problem. The BA/AB pooled total
p50 ratio is **0.999057300** for materialized and **0.982161217** for owning.
The corresponding phase ratios are, respectively:

| Route | Clone | Ingress | Snapshot | Edit and commit | Publish | Total |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Materialized BA / AB | 1.004485 | 0.985814 | 0.999076 | 0.992205 | 1.000075 | 0.999057 |
| Owning BA / AB | 0.958833 | 0.981094 | 1.006197 | 0.988625 | 0.963622 | 0.982161 |

Whole-process maximum RSS has p50 **95,990 KiB** for materialized and
**96,166 KiB** for owning, a **+176 KiB / 0.183%** owning difference. The
ranges are 94,984–96,956 KiB and 95,080–96,988 KiB. This is a process envelope
including generation, warmups, guards, samples and artifact work; it supplies
no operation allocation or peak-live-memory result.

## Guards and preservation evidence

All 24 reports bind the same 17,017,139-byte source
(`61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757`) and
17,017,145-byte output
(`23d6d7b8dd433ff453307ee56485efadc34d6b18b455a6858d9f565c4b1b6cd9`). Every
measured sample reports 260 sink writes with a largest write of 65,536 bytes,
and every total is at least the sum of its five recorded phases. The final
source review binds the semantic verification call in the harness; the report
gate fields are proof-state markers, while the independent reducer binds each
sample's output digest, size, sink counters and source identity.

Both representative route outputs pass the independent 0909 census. Each has
445 members, exactly the eight selected slide XML members changed, and 437
members unchanged at raw component level. Physical and central member order,
archive comment and non-selected metadata remain equal; data descriptors are
present and ZIP64 is absent. The census explicitly limits this evidence to the
generated descriptor/data-descriptor fixture and does not infer writer internals
or broader ZIP preservation behavior.

## Custody and cleanup

The build and capture are bound to parent `468fe2931d6532c32f4c6608590816b8eaa88e31`,
the 9,355-file source manifest, the retained Rust 1.95.0 release binary
(`f4ad7ee158061cc80354c3c8674c9fe6a8100ed76a0c8f743653efcb06bb2e8d`), and
the frozen plan. The 24 receipts agree on source manifest, helper, plan,
binary, route schedule and CPU pinning; process windows are serial and
non-overlapping. Build, scoped format, and warnings-denied Clippy receipts all
exit zero.

`replay-before-cleanup.json` and `replay-after-cleanup.json` both pass with
9,355 source files, 72 capture artifacts, 24 legs, and byte-exact census
replays for both routes. Cleanup passes, verifies the 9,355-file inventory,
removes no worktree, leaves unrelated worktrees untouched, and removes
1,538,498,852 logical bytes from the owned source/build roots. The retained
binary and generated archives remain Gitignored local evidence with tracked
descriptors.

The result is suitable as same-build existing-route evidence for this generated
in-memory fixture. It does not justify a production provenance bridge or a
general speedup claim. The next useful measurement is attribution of the
owning route's opened snapshot phase, preserving the same sink, teardown,
semantic and raw census boundaries. Cold cache, filesystem/range input,
producer variation, concurrency, allocation counters, ZIP64 and adversarial
descriptor forms remain outside this result.

## Hash bindings

| Artifact | SHA-256 |
| --- | --- |
| Final report `docs/performance/0910-pptx-owning-input-native-control.md` | `3c155ffd3a97a2c9f9bb99eb8ceb21d2ef1d42bfc581f921d5a91e41c14ef5e6` |
| Canonical analysis `analysis.json` | `62f8dcd7ad6f226d1e5d3527bd6afdcce5ca12775d1a826bb5b6fda105f7acf2` |
| Analysis method | `0842bb233ed6b31f870dfebcccc5889ae34275ea074d82c9fea11934ee5370b6` |
| Frozen capture plan | `b4666619aba3b6bbfb68955e70c9d47cad909c25c11b4affc39d9bb4f61f0370` |
| Census run receipt | `9e124d3488623767fe1d43f68761ca8bcfa7eb567fc2eb9e5e1bab3db44642dc` |
| Owned census | `58843ec4d50107fae674c88baa634839d8f45c5afbfa0a00c50f511ce64be036` |
| Materialized census | `e0e581d8b9368c6980ac35a31e130bcb4a746688582c0e1d17654e73d730de18` |
| Replay before cleanup | `16ca4e64a891cd93defc80cc9dd55906b0a4a5636a40fd6c909358eb54e44257` |
| Replay after cleanup | `16ca4e64a891cd93defc80cc9dd55906b0a4a5636a40fd6c909358eb54e44257` |
| Cleanup receipt | `1bcf0e8756db10349a3f10eb2c01df1b04aaf26273463ee98305ccd74c302f03` |
| Binary identity receipt | `b551a9bdc5b3ca1d4f2f27efdc91252edab8c280f0705cf6a14bec45f6561748` |
| Source manifest | `7a6d1310fb97dde1f275ec8702efabbef2cfb643c9ee4c1178169fed53aa4586` |
| Final source review | `0ddf5d1d3f313e59d16b875fbae4f6b32d1753e8ee7f8d07be1582b59d385888` |
| Protocol review | `d527584ebeeb835e5a8db5c9bbe118f3658059de35c731da42fab129e12cf6aa` |
