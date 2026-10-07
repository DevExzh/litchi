# 0959 inline Number comparison protocol

This packet is a prospective comparison for the isolated inline `Number`
candidate. It does not adopt the change or make a production claim. The
candidate executable, source freeze, build receipts, helper bindings, host
receipt, and root GO must be added to the plan before capture.

## Roles and custody

Role A is the retained 0957 native or allocator executable. Role B is the
new candidate executable built from the candidate source freeze. The retained
baseline files are byte bound before capture:

| lane | executable | bytes | SHA-256 |
| --- | --- | ---: | --- |
| native | `/home/zhuhe/litchi-performance-0957/executables/xlsx_managed_compression_benchmark.v2` | 43,259,328 | `5886b0f86a697bae757994d5a757cffae4a7e8d61a2769ff1c967029cd1912c3` |
| allocator | `/home/zhuhe/litchi-performance-0957/executables/xlsx_managed_compression_alloc.v2` | 43,676,872 | `9ba37a6b9796ad0e316e7571ebe117ba7a5f97bcf5d29153c7cb3d74c485e9a5` |

The candidate build must use Rust 1.95.0, release `opt-level=3`, fat LTO,
one codegen unit, debug information level 1, incremental disabled, and jobs
2. The exact candidate command, environment, target directory, source
manifest, terminal, binary, and result receipts are supplied by the build
owner and bound in the final plan. No build runs with a workload, and no
build overlaps capture.

The 0959 baseline freeze is `source-baseline-v2` with
`source-manifest-probe-v2.json` and `receipts/source-baseline-v2.json`.
Its one-cell layout probe is an explicit test-only overlay and is excluded
from runtime variation. The retained binaries remain bound to the separate
0957 runtime source manifest and source root. The candidate freeze must
contain a complete source census and the reviewed allowlist
`crates/litchi-xlsx/src/cell.rs`,
`crates/litchi-xlsx/src/raw/worksheet/semantic.rs`, and
`crates/litchi-xlsx/src/raw/worksheet/selected.rs`; an extra file, symlink,
changed manifest, or unbound build input rejects the packet. The final plan
binds one exact candidate source manifest in both its candidate-change and
lane descriptors. The manifest diff against source-baseline-v2 must contain
exactly those three paths; generated Cargo.lock files are admitted only when
present with identical bytes and hashes in both manifests.

The candidate native and allocator build result receipts must each report a
passed status, a zero terminal return code, no timeout, the candidate source
manifest SHA, and a binary descriptor whose bytes and SHA match the lane
binding. A bound host/topology/governor receipt and root GO receipt are also
required before execution.

Each child is a fresh process launched from the external packet directory
`/home/zhuhe/litchi-performance-0959`, with workers `1` and affinity
`12-15`. The runner uses `/usr/bin/taskset`, `/usr/bin/time -v`, an explicit
`LANG=C`, `LC_ALL=C`, `PATH=/usr/bin:/bin`, `RUST_BACKTRACE=0`, and `TZ=UTC`
environment. It permits one child at a time. A child is launched once; a
timeout, process error, report error, oracle mismatch, or missing metric stops
the lane and retains the failure packet. There is no retry or replacement.

For every child the runner writes the frozen command and environment before
launch, verifies the live affinity, waits for the process, and seals
`terminal.json` before reading `time-v.txt` or `report.json`. Only then does
it validate the report and write `validation.json` and
`post-terminal-artifacts.json`. The outer terminal receipt records the number
of started children and remains failed if the schedule stops early. A normal
terminal also requires the entire session process group to clear; an
interruption or launch error is recorded as a failed child terminal after
bounded SIGTERM/SIGKILL cleanup and reaping.

## Cases and exact oracles

The same three finite-context cases run in both lanes and in every role:

| case | source bytes / SHA-256 | output bytes / SHA-256 | semantic SHA-256 | CpuTasks | updates |
| --- | --- | --- | --- | ---: | ---: |
| `medium-one-percent` | 4,226,429 / `dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036` | 4,226,788 / `fc742a7ad139e2ce2e75fabfff8d2a33408dd5eeb10997717a2e2656d8423e6c` | `f705d4200e8bbe840abb16ff7d17da887cbe1d0fb53a649c69279451ff5d829c` | 5 | 93 |
| `dense-sparse-one-percent` | 4,251,863 / `893ad3f5dd6a98aec44bc541a140048072c84c579b4b9e332431f779b097cb1a` | 4,252,436 / `07eee3f980949f031776087ed61f06125edd82cb29394a608a3d8d5a4b28374d` | `d8bbafac4c31288eb48a773df88b6b2add2a51b9ad6fb652ab1aeb801c5dfa8c` | 5 | 178 |
| `medium-one-edit` | 4,226,429 / `dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036` | 4,226,480 / `9b7b66a02007eeb63498fd5de4c6b7115ace0383ce37d97e1a9560ef7bfadec1` | `3cd21160d4f74fa0f097ab40be08e211b3e460cea788aa2b6705a55fdece07de` | 2 | 1 |

Reports must retain the finite context limits, exact output and semantic
hashes, eager output, actual managed-commit inverse preflight, per-sample
inverse restoration, sink capacity and 64 KiB maximum write, and CpuTasks.
Native rows must report zero `resources_after_drop`; allocator rows must
report zero `resources_after_core`, which is the resource snapshot exposed by
that binary. No postdrop resource fields are invented for allocator rows.
Workers must report one requested worker and zero expected/observed worker
permits for all three cases.

## Fresh-process paired schedule

The native block sequence is six blocks per case:

```text
block 0 ABBA, block 1 BAAB, block 2 ABBA,
block 3 BAAB, block 4 ABBA, block 5 BAAB
```

`A` is baseline and `B` is candidate. Adjacent positions `(0,1)` and
`(2,3)` are the paired legs inside each native block, giving 12 paired legs
per case and 72 fresh native children. The allocator lane uses three fresh
pairs per case in the order `AB`, `BA`, `AB`, giving 18 allocator children
total. The native lane is run to a sealed terminal packet before the
allocator lane starts. Allocation counters therefore cannot be confused with
native timing evidence.

Native children use 20 warmups and 200 measured samples. Native reports retain
`ingress_ns`, `edit_commit_ns`, `publication_ns`, `commit_drop_ns`, and
`total_ns`; `edit_commit_ns` and `total_ns` are primary, while publication is
a control and the other phases remain diagnostic. Allocator children use 2
warmups and 10 measured rows. Allocation
reports retain setup, the four core phases, worker drain, both combined core
views, and postdrop with the `serialized_region_peak_v3` counters. Their
worker drain is the fifth contiguous core segment; counter conservation and
adjacent live-byte boundaries are checked by the report producer and runner.

## Review gates

Ratios are candidate over baseline, so lower is better. Missing, nonfinite, or
zero-denominator values fail closed. The declared bootstrap uses 10,000
resamples with seed 959 over six native block-level geometric means or three
allocator paired repeats for each case and metric.

For each native primary p50 and p95 edit-commit and total metric, the case
aggregate CI upper endpoint must be at most `1.05`. Individual paired-leg or
block ratios above `1.05` remain visible as review flags; they are not
silently dropped or promoted to adoption gates. A named case and metric may
claim at least 5% improvement only when its aggregate point is at most `0.95`,
the two-sided bootstrap CI upper endpoint is below `1.0`, and the practical
effect check passes. Native ingress, publication, and commit-drop p50/p95/p99
remain reported; publication is the control, p99 is review-only, and neither
is pooled into the primary claim. The p95 comparison is the declared tail
gate.

`max_rss_kib` comes from `/usr/bin/time -v` and covers the whole process,
including setup and oracle work. Its case aggregate CI upper endpoint must be
at most `1.05`; individual pair/block ratios above `1.05` remain review
flags. Any RSS reduction claim also requires aggregate point `<=0.95` and CI
upper `<1.0`, and is described as whole-process RSS evidence rather than
allocator ownership.

Allocator counters are descriptive, with no native latency inference. The
review retains `allocated_bytes`, `allocation_calls`,
`region_peak_live_bytes`, `core_combined_with_drain.allocated_bytes`, and
whole-process RSS as review metrics, while preserving every phase counter.
Each named allocator metric is subject to the same case-aggregate `1.05`
safety gate; a 5% reduction claim uses aggregate point `<=0.95` and CI upper
`<1.0`. A failed
control (`dense-sparse-one-percent` or `medium-one-edit`) prevents a broad
candidate claim. Any correctness, source census, binary, affinity, worker,
resource, conservation, terminal, or report-schema failure rejects the
comparison regardless of timing ratios.

## Commands and authorization boundary

The external helper is
`/home/zhuhe/litchi-performance-0959/tools/capture_0959_v4.py`.
Before root binds the candidate and GO, these commands perform only static
schedule checks:

```sh
python3 /home/zhuhe/litchi-performance-0959/tools/capture_0959_v4.py \
  --plan /home/zhuhe/litchi-performance-0959/plan-template-0959-v4.json \
  --lane native --dry-run
python3 /home/zhuhe/litchi-performance-0959/tools/capture_0959_v4.py \
  --plan /home/zhuhe/litchi-performance-0959/plan-template-0959-v4.json \
  --lane allocator --dry-run
python3 /home/zhuhe/litchi-performance-0959/tools/controls_0959_v4.py
```

The controls must report four rejected mutations before root GO: native role
order, allocator child count, native CPU mask, and a case output oracle. The
control summary and its runner/plan descriptors are retained and bound under
`helpers.controls`; they execute only dry-run plan validation.

After the final plan is status `authorized`, binds both candidate binaries,
candidate source/build receipts, capture/reducer/control helpers, and root
GO, the root-owned launch commands are:

```sh
LITCHI_PERF_0959_CAPTURE_APPROVED=1 \
python3 /home/zhuhe/litchi-performance-0959/tools/capture_0959_v4.py \
  --plan /home/zhuhe/litchi-performance-0959/plan-final-0959-v4.json \
  --lane native --execute \
  --packet /home/zhuhe/litchi-performance-0959/capture-native-v1

LITCHI_PERF_0959_CAPTURE_APPROVED=1 \
python3 /home/zhuhe/litchi-performance-0959/tools/capture_0959_v4.py \
  --plan /home/zhuhe/litchi-performance-0959/plan-final-0959-v4.json \
  --lane allocator --execute \
  --packet /home/zhuhe/litchi-performance-0959/capture-allocator-v1
```

The plan template and runner deliberately reject execution while the
candidate binding is pending. These commands are recorded for reproduction;
this protocol preparation has launched no workload and performed no build.
