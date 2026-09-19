# Change 0680 evidence: DOCX paragraph-index memo

Status: retained design evidence only. No production source was modified or
committed for this record; the probe build compiled the workspace production
dependencies into a task-specific external target, which is not retained.
performance_claim: none; the counts below are not timing, throughput, RSS, or
speedup claims. iWork is outside this change.

The design is [../../0680-docx-paragraph-memo-design.md](../../0680-docx-paragraph-memo-design.md).
The machine-readable decision record is [decision.json](decision.json).
The source audit is [measurements/source-audit.txt](measurements/source-audit.txt),
and the readable callgrind totals are
[measurements/callgrind-summary.txt](measurements/callgrind-summary.txt).
The reproducible probe source and lockfile are in [probe/](probe/).

## Reproduction

Environment captured for this packet:

~~~
Linux 7.0.0-1012-aws
rustc 1.95.0 (59807616e 2026-04-14)
cargo 1.95.0 (f2d3ce0bd 2026-03-21)
Valgrind-3.26.0
~~~

Revision and retained probe-source hashes:

~~~
HEAD 389167b38c8b43c84c0d1799369e17360a59209c
Cargo.toml 8897e4af659b8aadccf26bca6962095adc2f7c3f8bdfd1347315d75f0470608f
Cargo.lock 3d5e0c487282a2f93cba695ee63720ffab498985d56dff43a9939aa8b67998ee
src/main.rs e7bfcfe6faaa0d5ac0c19a758b23a314b345efa7dd3db8faeb0f34610dad860a
CPU AMD EPYC 9R45; 32 x86_64 CPUs; callgrind pinned to CPU 12
~~~

Build the retained source in a task-specific external target directory:

~~~
cargo build --manifest-path docs/performance/results/change-0680/probe/Cargo.toml \
  --release --offline -j2 --target-dir /home/zhuhe/code/litchi-target-0680-probe
~~~

The probe accepts operation paragraphs repetitions. The allocator sampling
commands run the release binary directly for each operation and each of
repetitions = 0, 1, 5; the package fixture is created and opened before the
sampled loop. The sampled line is the difference between counters immediately
before and after the loop. Every count operation asserts
document.paragraph_count() == paragraphs inside the loop.

The six operation names are:

~~~
eager_document
eager_document_count
eager_one_view_count
source_document
source_document_count
source_one_view_count
~~~

For instruction totals, run the same matrix under a pinned callgrind process,
for example:

~~~
taskset -c 12 valgrind --tool=callgrind --cache-sim=no --branch-sim=no \
  --callgrind-out-file=/tmp/change-0680.out \
  /home/zhuhe/code/litchi-target-0680-probe/release/probe0680 \
  eager_document_count 10000 5
~~~

For every operation and fixture, use the paired subtraction
sampled(reps) = total(reps) - total(reps=0). The 0-run is retained in the
summary because it proves package construction/open was outside the sampled
allocation region and makes the instruction subtraction reproducible. Raw
.out profiles and the external target are generated artifacts and are
intentionally not retained. The small .log files are retained because they
contain the probe command and sampled allocator/diagnostic lines.

## Results

The generated fixtures have these exact dimensions:

| paragraphs | visible XML bytes | archive bytes |
| ---: | ---: | ---: |
| 200 | 10,113 | 1,454 |
| 10,000 | 500,113 | 27,080 |

The semantic delta is the fresh count row minus the fresh document row. It is
also reproduced by the first count on one reused view:

| paragraphs | route | fresh view index allocation | one-view first index allocation | fresh-view index instructions |
| ---: | --- | ---: | ---: | ---: |
| 200 | eager | 6,287 B | 6,287 B | 906,624 Ir (1 rep); 907,016 Ir/view (5 reps) |
| 200 | source-backed | 6,287 B | 6,287 B | 906,682 Ir (1 rep); 906,905 Ir/view (5 reps) |
| 10,000 | eager | 342,735 B | 342,735 B | 44,867,496 Ir (1 rep); 44,867,499 Ir/view (5 reps) |
| 10,000 | source-backed | 342,735 B | 342,735 B | 44,883,417 Ir (1 rep); 44,883,701 Ir/view (5 reps) |

The one-view repetitions = 5 allocation delta remains the one-build value, not
five builds: 6,287 B for 200 paragraphs and 342,735 B for 10,000. Its later
callgrind totals are within the process/profile noise of the first build; the
probe uses the allocator and semantic assertions as the decisive repeated-work
witness.

The source route's diagnostics are cold_loads = 1, hits = 4,
successful_loads = 1, retained_entries = 1, with retained_bytes = 10,113 or
500,113 respectively. Thus the physical source payload is already shared on
the repeated-view path; the repeated operation is the DOCX semantic
index/view construction.

The complete paired totals are in callgrind-summary.txt; each sampled
allocation delta and source diagnostic is also retained in the matching .log
files under probe/measurements/callgrind/. No claim-registry entry is made
from this packet.

Retained terminal logs normalize trailing whitespace only; command lines,
allocator counters and instruction counts are unchanged.
