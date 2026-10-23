# 0362ye fixed admission probe

This is reduced allocator evidence for the fixed-buffer ODP admission probe.
The source commit is the commit that adds this receipt. The generated archive
is the tracked `odp_0362ye_probe` fixture (2,000 slides, 4 MiB media).

| route | e4a baseline peak / allocated | current peak / allocated | output bytes | output SipHash |
|---|---:|---:|---:|---:|
| `open` | 693231 / 777196 | 83233 / 167605 | 0 | 15130871412783076140 |
| `first_slide` | 693231 / 3210784 | 89988 / 2681513 | 47 | 15128627771916252792 |
| `first_text` | 1500637 / 4672388 | 990939 / 4143117 | 105779 | 7807713202416953651 |
| `unchanged_write` | 6074765 / 6782817 | 5484747 / 6173226 | 4215428 | 12221861870545843138 |
| `append_commit_write` | 612845 / 5737255 | 612893 / 5853815 | 4215645 | 5546059324926701325 |
| `replay` | 612845 / 5737255 | 612893 / 5853831 | 4215645 | 5546059324926701325 |

The earlier fixed-buffer candidate binary is kept outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/odp-0362yd-candidate-target/release/examples/odp_0362ye_probe`
(3,599,992 bytes, SHA-256
`032047b9f559c291cc07b1052e5425b688a2e91784220c7ef57672d091136bf0`).
The post-lint source build used for the current Callgrind rows is
`/home/zhuhe/code/litchi-worktrees/uo2-odp-heap-repair/target/release/examples/odp_0362ye_probe`
(3,599,992 bytes, SHA-256
`f4b3e47bdcceb9e9e536a3c94b181bde040360ffd66c78ead73feb74562737f9`).
The baseline binary is
`/home/zhuhe/code/litchi-worktrees/.coord/odp-0362yd-baseline-target/release/examples/odp_0362ye_probe`
(3,096,104 bytes, SHA-256
`ecb034b0f57588c7d131ba9981d7004f5181ee214a88d14519f73660f35e0669`).

The candidate was built and run with:

```text
ODP_PROBE_ROUTE=docs/unified-ops/results/0362ye/route_after.rs \
CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=0 \
CARGO_TARGET_DIR=$CANDIDATE_TARGET TMPDIR=$ODP_TMP \
cargo build --offline -p litchi-odp --example odp_0362ye_probe --release
$CANDIDATE_TARGET/release/examples/odp_0362ye_probe open
```

The baseline used the same command with `route_before.rs` and its protected
target. This is allocation evidence; the route-local Callgrind comparison
below uses the same fixed baseline binary and the current post-lint candidate.

## Route-local Callgrind

The fixed baseline rows are from
`/home/zhuhe/code/litchi-worktrees/.coord/odp-0362ye-evidence/callgrind-4b6e2c23bc`;
the current candidate rows are in
`/home/zhuhe/code/litchi-worktrees/.coord/odp-0362ye-evidence/callgrind-5575-candidate`.
Both use `run_callgrind.sh`, one process per route, and the same ten route
arguments. The captures remain outside the repository.

| route | baseline Ir | candidate Ir | ratio |
|---|---:|---:|---:|
| `open` | 829,907,975 | 829,845,468 | 0.999925× |
| `first_slide` | 886,558,187 | 894,428,125 | 1.008877× |
| `repeat_slide` | 954,665,843 | 970,662,419 | 1.016756× |
| `first_text` | 899,838,108 | 907,901,763 | 1.008961× |
| `repeat_text` | 969,740,535 | 985,835,146 | 1.016597× |
| `write_text` | 877,897,998 | 889,197,163 | 1.012871× |
| `repeat_write_text` | 925,853,231 | 948,745,773 | 1.024726× |
| `unchanged_write` | 845,074,834 | 845,016,713 | 0.999931× |
| `append_commit_write` | 1,349,361,627 | 1,351,626,925 | 1.001679× |
| `replay` | 1,349,364,328 | 1,351,598,499 | 1.001656× |

Native timing remains pending a quiet host window; these deterministic
instruction rows do not make a native-speed claim.
