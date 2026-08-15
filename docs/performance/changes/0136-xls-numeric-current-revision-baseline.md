# Change 0136: native XLS fixed-width numeric current-revision baseline

Date: 2026-08-15

## Scope

This record binds the four selectors introduced by change 0135 to a pinned
release baseline before any plan-only or lazy-target publication work begins.
It is a descriptive current-revision baseline, not a before/after result and
not an accepted optimization claim.

The measured revision was clean `9577cd16f36b0fd76be4eb1b1842278b4d2cf192`.
The exact release binary SHA-256 was
`f745bd905705036261a3a3a2e3953483e790a3afa558e296bc9c16eab41cd596`.
The raw schema-1 artifact is
[`xls-numeric-before-9577cd16f.json`](../results/xls-numeric-before-9577cd16f.json),
SHA-256
`2bd79c4b194584834cf5bfdf40ebc997631e7b360ef6a954358469f62179d073`.

## Environment and command

The run used Linux 6.8.0-101-generic, Rust 1.95.0, the system allocator, an
AMD EPYC 9575F host, and affinity-visible CPU 2 only. Each case used 20 warmup
iterations and 200 measured iterations in one release process:

```sh
git worktree add --detach /tmp/litchi-xls-baseline-clean-0136 \
  9577cd16f36b0fd76be4eb1b1842278b4d2cf192
CARGO_TARGET_DIR=/tmp/litchi-xls-baseline-target-0136 \
  cargo build --release --locked \
  --manifest-path /tmp/litchi-xls-baseline-clean-0136/tools/perf-baseline/Cargo.toml
cd /tmp/litchi-xls-baseline-clean-0136
taskset -c 2 /tmp/litchi-xls-baseline-target-0136/release/litchi-perf-baseline \
  --warmup 20 --samples 200 \
  --case xls_numeric_eager_number_edit_save,xls_numeric_source_backed_number_edit_save,\
xls_numeric_eager_rk_mulrk_edit_save,xls_numeric_source_backed_rk_mulrk_edit_save \
  --json /home/zhuhe/CodeProjects/litchi/docs/performance/results/xls-numeric-before-9577cd16f.json
```

The executable was built and run from a clean detached checkout of the measured
revision, using the dedicated target directory shown above; the temporary
worktree and target were removed after validation. The binary hash is the
authoritative executable identity. No cold-cache, fresh-process-per-sample,
allocation, peak-heap, RSS, hardware-counter, or physical-I/O evidence is
present in this artifact.

## Results

All values are milliseconds except the artifact-size column.

| Selector | p50 | p95 | p99 | mean | commit p50 | publication p50 | complete target retained |
|---|---:|---:|---:|---:|---:|---:|---:|
| eager Number | 31.492 | 34.116 | 35.916 | 31.763 | 30.741 | 0.729 | 16,995,840 B |
| source-backed Number | 146.410 | 149.108 | 150.693 | 146.642 | 101.618 | 44.783 | 16,995,840 B |
| eager RK/MulRK | 0.100 | 0.120 | 0.127 | 0.103 | 0.097 | 0.003 | 202,752 B |
| source-backed RK/MulRK | 1.627 | 1.659 | 1.690 | 1.630 | 1.117 | 0.509 | 202,752 B |

The source-backed/eager p50 ratios are 4.65x for Number and 16.25x for
RK/MulRK. These are matched implementation baselines that produce byte-identical
outputs within each family; they are not a before/after regression
classification. The separately sorted commit and publication phase medians
are descriptive and need not add exactly to the median of the per-sample phase
sums.

## Correctness and interpretation

Every case retained 200 samples with aligned phase vectors, the eager/source output SHA-256 values
matched within each family, and all samples reported complete target
materialization. Source ingress, expected-output construction, complete
Snapshot/Workbook reopen, CFB directory and untouched-stream checks,
patch/inverse/stale/no-op checks, typed security refusals, partial-sink checks,
and the real-producer `54016.xls` gate remained outside timing.

The evidence confirms that source-backed fixed-width publication still pays
for a complete target snapshot and substantial commit/publication work. It does
not prove which internal substage should be removed, nor does it establish an
allocation or memory benefit. Any plan-only candidate must retain complete
composed-target validation, public numeric readback, exact source and target
fingerprints, partial-output behavior, and every preservation/security gate,
then be compared through a balanced pinned release run.
