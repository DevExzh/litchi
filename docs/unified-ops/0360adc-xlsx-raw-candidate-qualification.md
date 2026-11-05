# 0360adc — XLSX raw identity candidate qualification

This is a source-frozen candidate diagnostic for the current assembly at
`ff86f4a1c07f05a98a8dddf272daef56942656ea`. It is not a native timing claim
and does not compare the candidate with an optimized or otherwise different
baseline. The pinned historical result remains the separate provisional
record in `0360adc-xlsx-raw-qualification.md`.

## Reproduced source and build

The standalone probe and lock are the same pair used by the historical lane:

| item | SHA-256 |
| --- | --- |
| `tools/perf-baseline/unified-xlsx-raw-identity/Cargo.toml` | `30ff248844c9e21436fc3913baa4c4d363daffdc617e5e7d03404423410d5c70` |
| committed `Cargo.lock` | `8b61b3c5e17b1c32d89139c97a26145117794afae83dd8fef3c9a7aff8765ee6` |
| `tools/perf-baseline/src/xlsx_raw_identity.rs` | `a251f1199b5ba442dfd9664806b81f29c7832d071df66fc2d2d5f2ada6c97fd4` |
| `tools/perf-baseline/src/allocation_metrics.rs` | `0fcec3e5c4972031b23542fd359a38401e25f32a133c77ef4c6e75facf0710d5` |

The detached qualification worktree was
`/home/zhuhe/code/litchi-worktrees/uo2-xlsx-final-qualification` at the
candidate revision. The allocator build began from the same clean revision
in the integration worktree before it was frozen; the native build and all
diagnostics used the detached worktree. Both binaries therefore use the
same source revision, manifest, lock, probe, and feature policy. The default
feature selects the required-limit OPC constructor.

Both builds used one Cargo job, disabled incremental compilation and release
debug information, and a 20 GiB cargo scope. The approved pre-existing target
directories were reused without cleaning or overwriting their recorded
binaries:

```sh
source (deleted-coord-artifact)
CARGO_TARGET_DIR=(deleted-coord-artifact) \
  CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=0 \
  LITCHI_MEM_MAX=20G cargo build --release --locked \
  --manifest-path tools/perf-baseline/unified-xlsx-raw-identity/Cargo.toml \
  --bin xlsx-raw-identity-alloc --features allocator-metrics

CARGO_TARGET_DIR=(deleted-coord-artifact) \
  CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=0 \
  LITCHI_MEM_MAX=20G cargo build --release --locked \
  --manifest-path tools/perf-baseline/unified-xlsx-raw-identity/Cargo.toml \
  --bin xlsx-raw-identity-native
```

The resulting binaries were:

| binary | bytes | SHA-256 |
| --- | ---: | --- |
| `xlsx-raw-identity-alloc` | 5,403,176 | `95dee63e89a0874a284743356246cbb102687f6f76df30967d8290b8ea245a15` |
| `xlsx-raw-identity-native` | 5,395,808 | `c42457745b2fd971bd8e7dba0867db6f946d3b9a50a2ed9b636f1a298d18704f` |

## Fidelity and costs

Each report used five iterations and two warmups. Four deterministic corpora
(small/large multipart × one/eight caches) and four operations per corpus
give 16 operation rows. Both allocator and native reports had 16/16 rows,
zero failed correctness booleans, and at least one native relationship
identity per operation. The Callgrind report independently repeated all
16 rows with one warmup and one measured iteration.

The following are candidate-only medians. Times are nanoseconds from the
native report; the allocation column is the allocator report's
`region_peak_live_bytes`. Owned-open and first-selected-read are shown
separately because the latter is the deferred decode path.

| corpus | owned open ns | first read ns | open peak bytes | first-read peak bytes |
| --- | ---: | ---: | ---: | ---: |
| `small_single` | 36,201 | 3,150 | 102,206 | 100,624 |
| `large_multipart_single` | 80,980 | 5,630 | 284,166 | 282,584 |
| `small_multi_cache` | 38,161 | 3,360 | 119,358 | 117,776 |
| `large_multipart_multi_cache` | 82,701 | 5,680 | 301,910 | 300,328 |

Commit operation medians (native elapsed ns; allocator peak bytes) were:

| corpus | slicer insert | timeline insert | slicer replace | timeline replace |
| --- | --- | --- | --- | --- |
| `small_single` | 74,010 / 72,481 | 115,550 / 77,284 | 48,870 / 151,044 | 113,480 / 157,500 |
| `large_multipart_single` | 1,684,610 / 17,372,333 | 1,665,320 / 17,379,321 | 1,593,850 / 17,382,677 | 1,652,190 / 17,384,568 |
| `small_multi_cache` | 930,846 / 135,298 | 428,473 / 141,706 | 47,470 / 169,876 | 113,881 / 176,332 |
| `large_multipart_multi_cache` | 2,982,478 / 17,409,055 | 2,024,423 / 17,428,441 | 1,667,270 / 17,402,102 | 1,652,190 / 17,403,991 |

The deferred decode counters were captured before publication and after each
publication. They remained equal for insertion rows; replacement rows showed
the expected decoded-part increase. Large multipart rows therefore include
the raw aggregate metadata cost in both owned-open/first-read and commit
measurements rather than hiding it in setup.

## External evidence and Callgrind status

Raw evidence is outside the checkout at
`(deleted-coord-artifact)`.
The reduced report hashes are:

| artifact | bytes | SHA-256 |
| --- | ---: | --- |
| allocator fidelity JSON | 96,837 | `105fb5351b183803867f74d0178b0134ffbc1785e1b7e45bb69aac35bde91631` |
| native fidelity JSON | 40,263 | `4dd55e764a771fe2cdf41bcc4e1309563017f57ea4dca27e1346ea65face9241` |
| Callgrind JSON | 26,544 | `333298f1d421678f1ba852f5ec4abebec5bd82ce4e6ef8a311281ff8c2908d3c` |
| Callgrind capture | 394,758 | `65e852afe56c52e5a38906bba513454ecf3fcaa8064ddf7721815c398a3eb803` |

The Callgrind capture reported `3,600,365,466` instructions (`Ir`). The
historical capture reported `3,589,721,324` Ir, a descriptive diagnostic
difference only; neither is a native timing or acceptance conclusion.
Callgrind ran while the coordinator's workspace cargo/rustc checks were
active, and the before/after process snapshots are retained beside the
capture. Consequently this record makes no quiet-host claim.

The reproducible diagnostic command was:

```sh
taskset -c 0 valgrind --tool=callgrind --quiet \
  --callgrind-out-file=(deleted-coord-artifact) \
  (deleted-coord-artifact) \
  xlsx-raw-identity --api-variant candidate \
  --source-revision ff86f4a1c07f05a98a8dddf272daef56942656ea \
  --iterations 1 --warmups 1 \
  --json (deleted-coord-artifact)
```
