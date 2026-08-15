# Change 0138: balanced release evidence for plan-only native XLS numeric publication

## Decision

Change 0138 is the acceptance-grade measurement for the two opt-in plan-only
selectors added by [change 0137](0137-xls-numeric-plan-only-publication.md):

- `xls_numeric_plan_only_number_edit_save`
- `xls_numeric_plan_only_rk_mulrk_edit_save`

The plan-only path is accepted for the measured total-latency improvement in
these two deterministic fixed-width native XLS families. The result is not an
operation-only allocation bound or physical-I/O claim. A process-level
maximum-RSS reduction is accepted for the Number family only; the RK/MulRK RSS
directions disagree. Valid heaptrack profiles show lower whole-process
allocation totals but identical peak heaps in the sampled A/B pairs, so no
operation-only allocation or peak-heap improvement is accepted.

## Clean release provenance and environment

The release binary was built from a clean detached worktree at the exact
measured implementation revision:

```text
revision: da3bfb8ced98e71cb38c602079ad69e64a96cd2d
binary: /tmp/litchi-xls-plan-only-target/release/litchi-perf-baseline
binary SHA-256: c79814cb4cc6420a6c56666737466c147ce08e5a564d82b3fbc245bda9ee8c4b
binary bytes: 39,324,832
```

The build used `cargo build --release --locked` with a dedicated temporary
target. The harness recorded `git_worktree_dirty: false`, CPU affinity `2`,
AMD EPYC 9575F, Linux 6.8.0-101-generic, Rust 1.95.0 and the Rust system
allocator. The root worktree's unrelated protected dirty files were not part
of the build or run. The detached worktree and target were removed after the
artifacts were frozen; cleanup proof is recorded below.

The latency command was run as one process per leg:

```text
taskset -c 2 litchi-perf-baseline --warmup 20 --samples 200 --case SELECTOR --json RESULT.json
```

Each family used the strictly sequential order `A1, B1, B2, A2`, where A is
ordinary `SourceBackedCommit` and B is the forward-only plan-only commit. No
benchmark processes overlapped. The matched process-RSS command used the same
binary and selectors with `--warmup 3 --samples 30` under
`/usr/bin/time -v`, in the same order and with one process at a time. The raw
schema-1 JSON and `/usr/bin/time -v` records are listed in the artifact
manifest.

The Number corpus is the existing `Untouched!E21` 42 -> 43 corpus:

```text
archive: 16,995,840 B, SHA-256 6a57231ba681bc7dd38d447ebd5348ef3b1fefedeefb1e61c97f22faa074e53
Workbook payload: 80,946 B, SHA-256 c78e03ba3743132e04b08bf6f4579ceb1c112a22c441c1e036381d3e06c6d041
```

The packed corpus contains one RK and one two-cell MulRK record:

```text
archive: 202,752 B, SHA-256 61a649b081c24821b02aa5e69b6ad1dc53b0232019d3668dd3776f402989c594
Workbook payload: 1,665 B, SHA-256 1b57f77d776cc8d0ed0f5154f7f7db2abca7f5fa7e23ecc69dd316c1c0b65967
```

## Latency results

All values below are nanoseconds converted to milliseconds from the complete
`elapsed_ns` vector (edit + set + commit + publication). Phase vectors were
also retained separately; phase p50s are not added to the total p50 because
each vector is independently reduced.

| Family | Leg | Implementation | p50 | p95 | p99 | mean |
|---|---|---|---:|---:|---:|---:|
| Number | A1 | source-backed | 145.381011 ms | 149.190069 ms | 151.659560 ms | 145.876149 ms |
| Number | B1 | plan-only | 105.303682 ms | 108.132468 ms | 109.966693 ms | 105.589081 ms |
| Number | B2 | plan-only | 104.917507 ms | 107.025599 ms | 108.650603 ms | 105.129215 ms |
| Number | A2 | source-backed | 146.901573 ms | 150.219494 ms | 151.851884 ms | 147.117888 ms |
| RK/MulRK | A1 | source-backed | 1.632588 ms | 1.683829 ms | 1.725347 ms | 1.638516 ms |
| RK/MulRK | B1 | plan-only | 1.226052 ms | 1.251684 ms | 1.257071 ms | 1.228661 ms |
| RK/MulRK | B2 | plan-only | 1.228266 ms | 1.247259 ms | 1.262829 ms | 1.230375 ms |
| RK/MulRK | A2 | source-backed | 1.628041 ms | 1.656662 ms | 1.672475 ms | 1.630296 ms |

The paired candidate/control ratios are B/A; the percentage is the reduction
`100 * (A - B) / A`:

| Family | Pair | p50 B/A | p95 B/A | p99 B/A | mean B/A | Direction |
|---|---|---:|---:|---:|---:|---|
| Number | A1 -> B1 | 0.724329 (-27.57%) | 0.724797 (-27.52%) | 0.725089 (-27.49%) | 0.723827 (-27.62%) | agree |
| Number | B2 -> A2 | 0.714203 (-28.58%) | 0.712461 (-28.75%) | 0.715504 (-28.45%) | 0.714592 (-28.54%) | agree |
| RK/MulRK | A1 -> B1 | 0.750987 (-24.90%) | 0.743356 (-25.66%) | 0.728590 (-27.14%) | 0.749862 (-25.01%) | agree |
| RK/MulRK | B2 -> A2 | 0.754444 (-24.56%) | 0.752875 (-24.71%) | 0.755066 (-24.49%) | 0.754695 (-24.53%) | agree |

The documented acceptance policy requires B to be lower than A for p50, p95,
p99 and mean in both paired directions. Both families satisfy that policy for
the complete operation. The commit phase also agrees in both directions:
Number improves 39.93% / 40.84% p50 and RK/MulRK improves 35.82% / 35.49%.
Publication is a separately timed sink interval and is near-neutral; its
individual p99 direction is not accepted as an optimization claim. The
accepted claim is therefore limited to the complete operation and its commit
phase, not to publication in isolation.

## Process RSS and heaptrack

`/usr/bin/time -v` maximum resident set size is a whole-process VmHWM
observation, not an operation-only allocation bound. The matched configuration
was three warmups and 30 measured samples per leg:

| Family | A1 RSS KiB | B1 RSS KiB | B2 RSS KiB | A2 RSS KiB | Paired B/A | Direction |
|---|---:|---:|---:|---:|---:|---|
| Number | 216,276 | 193,072 | 193,448 | 216,528 | 0.892711 / 0.893409 | agree; process-level -10.73% / -10.66% |
| RK/MulRK | 142,460 | 142,848 | 142,336 | 142,716 | 1.002724 / 0.997337 | disagree; no RSS improvement accepted |

Heaptrack 1.5.0 was available and captured one A and one B profile per
family, sequentially, with the same three-warmup/30-sample configuration. The
correct invocation pinned heaptrack itself and launched the binary directly:
`taskset -c 2 heaptrack -o PROFILE litchi-perf-baseline ...`. The earlier
invalid profiles that tracked only the `taskset` wrapper were discarded and
are not part of the artifact manifest.

The valid whole-process profiles report the following:

| Family | Leg | allocation calls | temporary allocations | peak heap | peak RSS (including heaptrack overhead) |
|---|---|---:|---:|---:|---:|
| Number | A | 2,471,535 | 64,846 | 205.56 MiB | 225.96 MiB |
| Number | B | 1,576,870 | 41,179 | 205.56 MiB | 202.34 MiB |
| RK/MulRK | A | 472,976 | 45,050 | 154.93 MiB | 133.30 MiB |
| RK/MulRK | B | 306,251 | 28,183 | 154.93 MiB | 133.21 MiB |

These whole-process allocation and profiler-RSS totals are descriptive
single A/B profiles, not operation-only attribution or an ABBA resource
acceptance. Peak heap is unchanged in both families; therefore no peak-heap
improvement is accepted. The exact compressed profiles are retained in the
artifact manifest.

## Correctness and limitations

All eight latency JSONs and eight RSS JSONs are schema 1, identify the exact
revision, report a clean detached runtime worktree and CPU affinity `2`, and
pass their 200-sample or 30-sample vector-length invariants. Every leg emits
the same family output digest: Number
`f8f37064dc842550445b674385c196640c07681465de558b74dd2480b040fc03`, and
RK/MulRK
`ddf5d5b81d677f9056e5b48815c134e36d6674647f4e5f8c9946f989f39cf260`.
The plan-only legs retain zero complete target-artifact bytes and false
target-retention/materialization and patch/inverse flags; their sink bytes
still equal the complete CFB output. The existing focused/full correctness,
security, topology, partial-sink, no-op/fingerprint, forward-reopen and
54016.xls producer gates remain green from change 0137.

This record does not claim physical I/O, cold-cache behavior, syscall counts,
bounded total memory, operation-only allocation, hardware-counter behavior,
or broad native XLS producer coverage. It is limited to the two deterministic
fixed-width families and the recorded release binary/configuration.

## Frozen artifact manifest and cleanup

Latency JSONs are `xls-numeric-plan-only-0138-{number,rk-mulrk}-{a1,b1,b2,a2}.json`;
RSS JSONs and `/usr/bin/time -v` records use the same names with the `rss-`
prefix. Heaptrack profiles are the four `heaptrack-*.dat.gz.zst` files with
the 0138 prefix. Their SHA-256 values are recorded in
[`xls-numeric-plan-only-0138.sha256`](../results/xls-numeric-plan-only-0138.sha256).

The raw latency legs are [Number A1](../results/xls-numeric-plan-only-0138-number-a1.json),
[Number B1](../results/xls-numeric-plan-only-0138-number-b1.json),
[Number B2](../results/xls-numeric-plan-only-0138-number-b2.json),
[Number A2](../results/xls-numeric-plan-only-0138-number-a2.json),
[RK/MulRK A1](../results/xls-numeric-plan-only-0138-rk-mulrk-a1.json),
[RK/MulRK B1](../results/xls-numeric-plan-only-0138-rk-mulrk-b1.json),
[RK/MulRK B2](../results/xls-numeric-plan-only-0138-rk-mulrk-b2.json) and
[RK/MulRK A2](../results/xls-numeric-plan-only-0138-rk-mulrk-a2.json).
The matched RSS JSONs are [Number A1](../results/xls-numeric-plan-only-0138-rss-number-a1.json),
[Number B1](../results/xls-numeric-plan-only-0138-rss-number-b1.json),
[Number B2](../results/xls-numeric-plan-only-0138-rss-number-b2.json),
[Number A2](../results/xls-numeric-plan-only-0138-rss-number-a2.json),
[RK/MulRK A1](../results/xls-numeric-plan-only-0138-rss-rk-mulrk-a1.json),
[RK/MulRK B1](../results/xls-numeric-plan-only-0138-rss-rk-mulrk-b1.json),
[RK/MulRK B2](../results/xls-numeric-plan-only-0138-rss-rk-mulrk-b2.json) and
[RK/MulRK A2](../results/xls-numeric-plan-only-0138-rss-rk-mulrk-a2.json).
The corresponding `/usr/bin/time -v` logs are [Number A1](../results/xls-numeric-plan-only-0138-rss-number-a1.time.txt),
[Number B1](../results/xls-numeric-plan-only-0138-rss-number-b1.time.txt),
[Number B2](../results/xls-numeric-plan-only-0138-rss-number-b2.time.txt),
[Number A2](../results/xls-numeric-plan-only-0138-rss-number-a2.time.txt),
[RK/MulRK A1](../results/xls-numeric-plan-only-0138-rss-rk-mulrk-a1.time.txt),
[RK/MulRK B1](../results/xls-numeric-plan-only-0138-rss-rk-mulrk-b1.time.txt),
[RK/MulRK B2](../results/xls-numeric-plan-only-0138-rss-rk-mulrk-b2.time.txt) and
[RK/MulRK A2](../results/xls-numeric-plan-only-0138-rss-rk-mulrk-a2.time.txt).
The four valid heaptrack profiles are [Number A](../results/xls-numeric-plan-only-0138-heaptrack-number-a.dat.gz.zst),
[Number B](../results/xls-numeric-plan-only-0138-heaptrack-number-b.dat.gz.zst),
[RK/MulRK A](../results/xls-numeric-plan-only-0138-heaptrack-rk-mulrk-a.dat.gz.zst)
and [RK/MulRK B](../results/xls-numeric-plan-only-0138-heaptrack-rk-mulrk-b.dat.gz.zst).

After the artifacts were frozen, the detached worktree
`/tmp/litchi-xls-plan-only-0138`, dedicated target
`/tmp/litchi-xls-plan-only-target` and run-only profiler directory
`/tmp/litchi-xls-plan-only-runs` were removed, `git worktree prune` was run,
and the repository was checked to have one worktree, only the existing feature
and `main` branches, no stash entries and no benchmark process. The root
protected dirty files were preserved and no production or harness source was
changed by this measurement.
