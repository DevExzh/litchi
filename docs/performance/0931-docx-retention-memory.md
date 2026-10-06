# 0931 — DOCX retained-base allocation lifetime diagnostic

This batch adds an allocator-instrumented harness for the bounded base
inventory retention adopted in [0930](0930-docx-commit-base-retention.md).
Production behavior is unchanged. All 96 formal children pass. With eight
large-document commits, releasing optional bases frees 28,639,872 allocator
bytes for independent sources and 3,579,984 bytes for shared snapshots, despite
the same 28,641,536-byte logical sum. Every child returns to its measured
baseline after commit drop. This does not measure normal-release latency or
qualify a speedup.

## Workload and ownership

The deterministic minimal OPC documents contain 16 or 8,192 sections. Each
child holds one or eight distinct section-layout commits, using default
(1 MiB), zero, exact-weight, or one-byte-under-weight retention ceilings.
A finite 64 MiB probe obtains the public logical weight and verifies all
four admission boundaries. The two ownership modes use either clones of one
queried snapshot or independently opened source copies and snapshots.
Public inventory sharing checks verify the expected topology; commits are
created separately rather than cloned.

The commit count measures simultaneous ownership; construction and release
are sequential. This is not a multithreaded scaling measurement.

The prospective matrix has 32 cells and three fresh processes per cell,
pinned to CPU 12. All children follow the same journal: creation and external
snapshot release; correctness verification and package release; held commits;
release of the first N−1 optional bases; release of the last base; and commit
drop. Preallocated bookkeeping vectors retain their capacity through the last
checkpoint, allowing the final live-byte balance to return to its baseline.

The host reports an AMD EPYC 9R45, 32 visible CPUs, 132,553,801,728 bytes of
RAM, no swap, and Linux 7.0.0-1013-aws under KVM. The
[host record](results/change-0931/host-and-idle-v1.json) includes the observed
commands and outputs. The pre-capture process scan found no matching build,
test or benchmark jobs; it does not establish whole-host idleness.

Every distinct edit has a preflight exact-output, reopened-shape and inverse
check. Independent sources also validate every measured commit against its
original package. Since publication consumes a uniquely identified package,
the shared-source case publishes one representative measured commit against
that package and checks every measured target's shape. Its explicit validation
count is one; this is not evidence that every measured shared commit was
published.

## Measurement boundaries

The existing counting allocator records allocation, deallocation and
reallocation calls, requested byte totals, and live bytes. Creation and each
release/drop operation have separate region peaks. `commits_retained` is an
absolute checkpoint without a region peak. Correctness work runs outside
those regions, but contributes to the process-wide allocator high-water mark
and `/usr/bin/time -v` maximum RSS. Neither is a retention-only peak.

`Commit::retained_base_bytes()` is a logical per-commit admission weight.
Summing it can count shared inventory allocations repeatedly. It excludes
source XML already owned by the patch, required target state, inline commit
storage and allocator bookkeeping. Disabling optional retention therefore
does not make a commit allocation-free. Release deltas measure callback-visible
requested live bytes; physical resident pages need not fall at the same time.

The harness uses the existing isolated allocator wrapper and introduces no
production dependency or unsafe code. ADR 0003 snapshot/publication contracts,
ADR 0005 finite observable optional retention, and ADR 0031 execution-budget
boundaries remain unchanged. No new no-op measurement is included; 0930's
existing correctness coverage remains the reference for that case.

## Results

The frozen cohort completed once: 96 children, 32 cells, three repetitions per
cell, zero failures and no retries. The original outer tool exit was zero.
Allocator held/release live-byte deltas are identical across the three
repetitions in every cell. All admission, topology, shape, scoped output and
inverse checks pass; no observer-invalid or overflow flag is set.

The following exact-cap results hold eight commits. Values are bytes; live
values are relative to the pre-creation baseline. The release columns show
bytes freed by each operation, not remaining live bytes.

| Sections | Source ownership | Logical sum | Held live delta | First seven releases free | Last release frees | Live delta after all releases |
| ---: | --- | ---: | ---: | ---: | ---: | ---: |
| 16 | Shared snapshot | 58,240 | 85,197 | 0 | 7,072 | 78,125 |
| 16 | Independent sources | 58,240 | 151,312 | 49,504 | 7,072 | 94,736 |
| 8,192 | Shared snapshot | 28,641,536 | 42,755,052 | 0 | 3,579,984 | 39,175,068 |
| 8,192 | Independent sources | 28,641,536 | 76,008,832 | 25,059,888 | 3,579,984 | 47,368,960 |

Shared bases stay live until the final retaining commit releases them.
Independent bases free separately. The logical charge exceeds released heap
bytes by 208 bytes per independent base in these two shapes; accounting
includes an inline snapshot-owner charge, so equality with freed heap bytes
is not expected. One-commit cases free the same 7,072 or 3,579,984 bytes in
either ownership mode.

Default retention matches exact retention at 16 sections and zero retention
at 8,192 sections. Zero and one-under policies retain no optional base. Their
held live deltas match the corresponding exact-policy values after all bases
are released. Required commit targets and patch/source bytes still remain;
only commit drop returns the live-byte delta to zero.

Creation's requested bytes and region peaks above starting live bytes match
between exact and zero policies for each shape/topology. Absolute peaks differ
by three bytes of pre-existing process state. In this harness all queried base
snapshots remain live until the end of creation, so this observation does not show
that a zero cap reduces construction peak memory. At eight large commits,
creation requests 1,251,354,390 bytes with shared snapshots and 2,226,793,064
with independent sources; region peaks above the starting live bytes are
45,083,363 and 78,539,940 bytes respectively. Requested totals count repeated
allocations and reallocations, not simultaneously live memory.

Whole-process RSS remains descriptive and includes preflight and correctness
work. For the large eight-commit exact-cap cells, RSS ranges across three
processes are 76,492,800–77,238,272 bytes (shared) and
115,175,424–116,006,912 bytes (independent). These are different ownership
workloads, not a production before/after regression comparison. Release-region
peaks never exceed their starting live bytes; the held checkpoint has no
region peak. No claim of immediate resident-page reclamation follows.

The [full scalar reduction](results/change-0931/formal-reduction-v1.json)
retains every repetition and all 32 cells. The
[compact result](results/change-0931/comparison-memory-result-v1.json) binds
the original packet and copied replay. The generated DOCX sources are 1,052
and 25,498 compressed bytes; their hashes and output identities remain in
the reports and reduction. They are synthetic, highly compressible documents,
not representative real-world corpus coverage.

The cap and explicit release behave as intended for these shapes. There is
no evidence here to change the default ceiling. Next attribute the substantial
construction allocation traffic and profile the remaining publication work
on the adopted implementation, preserving mandatory source/output fingerprints.
This batch supplies a memory baseline rather than a production optimization.

## Quality and build history

The final harness passes two integration tests (small exact/shared and large
one-under/independent), two focused module tests, selected warning-denied
Clippy, warning-denied library rustdoc, and formatting for all four changed
Rust files. These are focused harness gates, not a fresh full production
workspace test run. Production code is unchanged from 0930.

Development attempts remain recorded in
[`new-quality.json`](results/change-0931/new-quality.json). A preliminary
shared-source publication attempt hit `SectionLayoutForeignSource`; the final
harness respects consuming package lineage and reports its representative
validation scope. That preliminary wrapper exit was not retained. The first
integration build produced a truncated linker failure and failed to write its
status file with `Disk quota exceeded`; the tool observed wrapper exit 101,
but Cargo's exit and the underlying linker cause are not established. Moving
the disposable target off tmpfs allowed the tests to finish. An unused test
import and three harness Clippy warnings were fixed before final checks.

The first measurement build command accidentally omitted `--release`. It was
interrupted with exit 130 before any workload ran, and its receipt remains
retained. The corrected build uses Rust 1.95.0, the standalone benchmark lock,
`allocator-metrics`, release optimization and explicit frame pointers. This
instrumented executable is dedicated to allocation diagnostics.

Two release preflights establish logical base weights of 7,280 bytes for 16
sections and 3,580,192 bytes for 8,192 sections. Both child exits and boundary
checks pass. Their wrapper exited 1 in a post-run display lookup after writing
the reports and terminal records; the children were not repeated. These two
runs are separate from the formal cohort. Review also corrected two ancillary
fixture paths in the source-closure addendum; the source and binary bytes did
not change.

## Reproduction

From this batch's source, build the diagnostic outside the checkout:

```sh
env -u CARGO_ENCODED_RUSTFLAGS RUSTFLAGS='-C force-frame-pointers=yes' \
  cargo +1.95.0 build --offline --locked -j2 --release \
  --manifest-path tools/perf-baseline/Cargo.toml \
  --features allocator-metrics --bin litchi-perf-baseline-alloc \
  --target-dir /home/zhuhe/litchi-performance-0931/reproduction-target
```

The allocator binary's `docx-section-retention` selector accepts `--sections`,
`--concurrent-commits`, `--policy`, `--ownership`, `--source-revision` and
`--output`. Use a fresh output file. The committed
[plan](results/change-0931/comparison-plan-v1.json) and
[runner](results/change-0931/comparison-0931-memory.py) define the complete
schedule, explicit child environment, source/binary bindings and terminal
records. Reproduction requires a newly bound manifest for the reproduced
binary; historical hashes cannot authorize a different build.

The retained raw packet archive is
`/home/zhuhe/litchi-performance-0931/formal-packet-v1.tar`: 7,680,000 bytes,
SHA-256 `faea4d1d0120feb07d69cb80bff155b1a446dce6f1bf33c50fb2403ec590e641`.
The [artifact receipt](results/change-0931/comparison-tooling-validation-v2.json)
records archive, source-list and reduction regeneration commands, sizes and
hashes. Original, relocated and independent replay reductions match SHA-256
`919e10175e0ee14cafd6e0acd7bfa5e28347d400cb3d6c0a9d75fa4ff25811aa`.
The [independent audit](results/change-0931/independent-postcapture-audit-v1.json)
records its separate copied-packet check.

[Cleanup](results/change-0931/cleanup-v1.json) removed 58,288,989,702 logical
bytes of disposable build intermediates and unused setup source copies.
Physical space reclaimed was not measured, and the logical sum may include
hardlinks. The final source, measured binary, raw evidence and receipts remain
outside the repository; the binary hash was rechecked after cleanup. The shared
checkout target and unrelated user changes were untouched. Every committed
evidence file is below 1 MiB, and no compiled executable is committed.
