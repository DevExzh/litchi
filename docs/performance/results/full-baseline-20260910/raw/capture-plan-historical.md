# Full baseline capture plan (prepared, timing deferred)

This artifact prepares the default release baseline and the checked allocator
tranche. It does not contain timing samples. The run script is gated by
`RUN_CAPTURE_AFTER_REVIEW=1` because the PPTX candidate is still under
correctness review.

## Source and lock identity

- Control source: `/var/tmp/litchi-performance-full-baseline-capture-control-20260910`
- Base revision: `1b3f2c2d0c059e8e59272775a97d0567e14e67f2`
- Frozen control remains at `/var/tmp/litchi-performance-full-baseline-control-20260909`; this plan never mutates it.
- The clean control lock SHA-256 is `8c731d7506bef3a4c05df553a0cd1cadb7d12da641099a2f9a0adc7e3fb4c518`.
- Build-only lock correction is `/var/tmp/litchi-performance-full-baseline-20260909/Cargo.lock.generated.patch`, SHA-256 `806b37859e566bbbf78f3f77ea0ecf1040ce5b09b0a9bc811b7f795c4117610b`. The script restores the clean lock before running binaries, allowing the reports to record a clean worktree.
- Expected normal identity is 201 rows with result-key digest `f0fd76293959e72211e06e51b0a2b41f371423fda34c55144077193d262b1670`.
- Expected allocator identity is two `opc_file_eager_open` rows (`warm`, `cold-requested`) with result-key digest `debb700930d4be1d2d597de60ef5b64200403407c138bab17dd268996975ec7a`; the checked catalog is `docs/performance/results/perf-regression-allocator-manifest-v1.json`.

## Build environment

The script uses Rust `1.95.0` with `RUSTUP_HOME=/tmp/litchi-spec-gap-rustup`,
`CARGO_BUILD_JOBS=4`, `CARGO_INCREMENTAL=0`,
`CARGO_PROFILE_RELEASE_DEBUG=0`, `RUSTFLAGS='-D warnings -D deprecated'`,
`LC_ALL=C.UTF-8`, and a dedicated target at
`/var/tmp/litchi-performance-full-baseline-20260910/capture-control/cargo-target`.
It builds the normal release binary and the separate
`allocator-metrics` release binary, then invokes the binaries directly so
Cargo is outside the measured process.

## Affinity and contention record

Both binaries run under `taskset --cpu-list 2`. CPU 2 is on-line and has one
logical CPU on this host. The harness records `environment.cpu_affinity` from
`/proc/self/status`; the wrapper additionally saves a pinned snapshot before
each report containing the timestamp, affinity, load average, online CPUs, and
top processes. The host is shared: the preparation snapshot recorded
loadavg `5.17 5.95 4.27` and active `rustdoc`, `git`, and harness processes,
so it is not an idle-host claim. Before the capture window, inspect each
`*-pre-run-contention.txt`; if CPU 2 is occupied by another benchmark, defer
and take a new snapshot. Keep the raw snapshots with the reports.

The default harness filesystem selection is explicit:
`--filesystem-cache warm,cold-requested`. These rows distinguish requested
cache state; `cold-requested` is not a proof of physical cache eviction.
No `--filesystem-root` is selected, so the harness records its default
filesystem evidence and does not claim storage isolation.

## Exact deferred invocation

From any directory, after the PPTX review clears:

```sh
RUN_CAPTURE_AFTER_REVIEW=1 \
  /var/tmp/litchi-performance-full-baseline-20260910/capture-control/run-full-baseline-and-allocator.sh
```

The script first applies the reviewed lock patch only for release builds,
restores the clean lock, and then executes these direct-binary commands on CPU
2:

```sh
taskset --cpu-list 2 env RUSTFLAGS='-D warnings -D deprecated' LC_ALL=C.UTF-8 \
  /var/tmp/litchi-performance-full-baseline-20260910/capture-control/cargo-target/release/litchi-perf-baseline \
  --warmup 3 --samples 15 --filesystem-cache warm,cold-requested \
  --json /var/tmp/litchi-performance-full-baseline-20260910/capture-control/full-normal.json \
  --corpus-manifest /var/tmp/litchi-performance-full-baseline-20260910/capture-control/full-normal.corpus-manifest-v2.json

taskset --cpu-list 2 env RUSTFLAGS='-D warnings -D deprecated' LC_ALL=C.UTF-8 \
  /var/tmp/litchi-performance-full-baseline-20260910/capture-control/cargo-target/release/litchi-perf-baseline-alloc \
  --warmup 3 --samples 15 --case opc_file_eager_open \
  --filesystem-cache warm,cold-requested \
  --json /var/tmp/litchi-performance-full-baseline-20260910/capture-control/opc-file-allocator.json \
  --corpus-manifest /var/tmp/litchi-performance-full-baseline-20260910/capture-control/opc-file-allocator.corpus-manifest-v2.json
```

The normal report is the 201-row wall-clock baseline. The allocator report is
allocation evidence only; its elapsed samples must not be used as latency
comparisons. The post-run checks assert the row counts, sample/warmup counts,
binary identity, allocator instrumentation identity, report hashes, and clean
control worktree. The run records no full-program completion claim by itself;
profile, I/O, peak RSS, scaling, ABBA, and other GOAL deliverables remain
separate evidence lanes.
