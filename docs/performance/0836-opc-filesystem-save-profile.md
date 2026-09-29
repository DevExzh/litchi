# 0836 — OPC source-backed filesystem save profile

This batch measures `opc_file_source_one_part_atomic_save` on the unchanged
source behind [0835](0835-filesystem-route-cache-baseline.md). It retains
34 reports / 444 measured outputs: two qualification reports / four outputs,
24 native controls / 288 outputs, four CPU profiles / 120 outputs, and four
syscall traces / 32 outputs. These populations are analyzed separately.
No production optimization or speedup is claimed; iWork is excluded.

The [packet](results/change-0836/README.md) retains source identities, command
receipts, raw reports, compressed perf records and decoded stacks, executable
mappings, compressed syscall traces, reader versions and validation failures.
The source revision is `5bb91a4de42a403cd1dfcb6e10058e3fe62d2228`.

## Scope and protocol

Both fresh release builds use the same 9,389 source inputs. The second adds
`-C force-frame-pointers=yes`. Workload captures run serially on CPU 12 of the recorded
AMD EPYC 9R45 host, Linux 7.0.0-1012-aws, Rust 1.95.0, with the system allocator
and recorded ext4 mount. CPU pinning does not reserve storage or CPU capacity.
The tiny perf-availability probe overlapped the frame-pointer build; it preceded
all workload captures and is the sole permitted command overlap.

The fixed OPC archive has four incompressible 4 MiB logical Parts and six ZIP
members. The ordinary source is 16,783,632 bytes, SHA-256
`a0c1af9e2c7a19148b44fc2a8c594c7a274131d74f9f042d55b487d5337cd1e6`.
One Part changes. Cold qualification adds the proven private 1,776-byte EOCD
comment and can add a bounded 64 KiB metadata-tail read. Cold residency and
process I/O observations do not establish physical-device reads.

Six native blocks alternate order across both builds and cache states. Every
command measures twelve fresh child processes after two warmups. Block p50
uses the integer midpoint; p95 and p99 both equal the block maximum with this
sample count. Reported latency summaries are medians of six block summaries.
Paired frame-pointer/baseline ratios use matched blocks, 10,000 bootstrap
resamples, seed 836083, and sorted endpoints 249 and 9,749.

| Cache state | Build | p50 ms | p95/p99 ms |
|---|---|---:|---:|
| Warm | Ordinary | 77.123311 | 80.023267 |
| Warm | Frame pointers | 79.638169 | 82.210855 |
| Verified cold | Ordinary | 144.921409 | 148.879293 |
| Verified cold | Frame pointers | 147.578952 | 150.518062 |

| Cache state | Metric | Paired fp/baseline ratio | Bootstrap 95% interval |
|---|---|---:|---:|
| Warm | Latency p50 | 1.031914 | 1.031417–1.034132 |
| Verified cold | Latency p50 | 1.018135 | 1.015426–1.019992 |
| Warm | Child peak RSS p50 | 1.023830 | 1.020792–1.030131 |
| Verified cold | Child peak RSS p50 | 1.029459 | 1.023885–1.035653 |

These ratios measure instrumentation effects, not an optimization. No planned
20% latency/RSS block-spread flag is triggered. Child peak RSS includes setup
and cold-preparation history; it is not peak live allocation above timer entry.

## Source and timer boundaries

The timed save includes package open, selected-target read/compare,
preservation planning, raw copying, target regeneration, buffered flush,
file synchronization, atomic rename, and parent-directory synchronization.
The selected original Part is decoded for exact no-op detection before the
changed Part is compressed at Deflate level 6. Untouched compressed members
are copied through in bounded chunks. Zero ordinary Part materializations
does not mean zero decompression. Near-archive-sized logical reads do not prove
a second complete archive read: the selected target is read for comparison,
and other member spans are copied. See the retained
[source review](results/change-0836/source-review.md).

The whole-save helper is inlined, so CPU admission uses the existing
`SourceBackedPackage::write_single_part_overlay_to_stream` function. Ten
emitted monomorphizations per build were admitted by raw/demangled symbols,
address/size agreement and bounded disassembly before capture. This scope
excludes package open and final atomic synchronization. Sampling uses inherited
`cycles:u` events at 997 Hz with frame-pointer stacks and no build-ID cache.
Measured child PIDs, exact executable paths and admitted owner ranges define
the selected population. Parent, primer and other samples remain separate.

Syscall traces use a separate ordinary build and `strace -f -qq -ttt -T -yy`
filtered to synchronization and rename calls. Their durations are perturbed by
ptrace and cannot be subtracted from native latency. Sampled CPU cycles do not
measure blocked time. Neither instrument provides an Amdahl wall-time fraction.

## Validation scope

Production Rust is unchanged. Historical quality evidence is reused with its
original scope: 569 tests passed and one was ignored before two helper
amendments; the final helper passed twelve focused tests and formatting,
compilation, Clippy, rustdoc and boundary gates. This batch does not claim a
new complete test-suite run. Both fresh builds, qualification and all native
and instrumented workload commands pass.

Preflight failures remain retained: a multi-sample reader test initially ran
before native reports existed; the CPU parser initially rejected `cycles:u`;
the independent audit assumed the wrong Cargo fingerprint filename and then
mistook result statistics for raw samples. An inferred PIE bias was removed
after an out-of-range mutation was accepted. The first trace-reader invocation
used the wrong relative manifest path. Cold traces additionally exposed a
pre-timer source `fsync`, requiring explicit path-based separation from atomic
publication. No failed workload is hidden or rerun to replace a measurement.

## CPU observations

The detailed reader agrees exactly with the independent mmap/ELF owner census
for all four captures. Build IDs, executable segment bounds, PID mappings and
sample timestamps bind runtime addresses to admitted static symbol ranges.
No load bias is inferred from the samples.

| Capture | Cache | All samples | Measured-PID samples | Owner samples | Deflate | Inflate | Other | Unknown stacks, whole capture |
|---|---|---:|---:|---:|---:|---:|---:|---:|
| 0 | warm | 26,429 | 11,662 | 1,886 | 1,853 | 1 | 32 | 1,080 |
| 0 | cold-verified | 27,934 | 12,858 | 1,863 | 1,801 | 4 | 58 | 1,125 |
| 1 | warm | 26,419 | 11,638 | 1,881 | 1,843 | 4 | 34 | 1,070 |
| 1 | cold-verified | 28,573 | 13,445 | 1,870 | 1,815 | 4 | 51 | 1,139 |

These are counts of observed samples, not wall-time percentages. The copy/CRC
name bucket is zero in every capture; that does not establish zero copying.
Every capture has zero explicit lost-event lines, malformed samples, empty
stacks and duplicate timestamps. Unknown frames remain: owner-interior unknown
counts are 12, 24, 8 and 25. Under the conservative reader policy, **CPU
fractions are withheld in all four captures**. Parent/primer samples and
measured-PID samples outside the overlay remain counted separately.

Deflate dominates the observed owner stacks. The largest self leaf is
`zlib_rs::deflate::algorithm::medium::deflate_medium` (1,292 / 1,246 / 1,307 /
1,260 samples), followed by `longest_match` (405 / 419 / 387 / 419). This
supports prioritizing changed-Part compression in the next measured experiment.
It does not authorize changing default preservation, compression or durability
policies, nor establish a whole-save CPU or wall fraction.

## Durability observations

The original whole-PID two-fsync assumption was too broad for cold children.
The amended offline scope verifies temp-file fsync, rename of that exact temp
path to the destination, then destination-parent fsync. All three calls must
succeed. Each cold child also has exactly one earlier fsync of the qualified
cold source; it is retained outside the publication scope. Warm children have
no excluded measured-PID event. No capture is replaced.

Each row below summarizes eight measured children. Component columns are
separate medians; the final duration is the median of per-child sums and need
not equal the sum of the component medians. All durations are ptrace-observed.

| Capture | Cache | File fsync ms | Rename ms | Parent fsync ms | Three-call sum p50 ms | Excluded preparation events | Other-PID events |
|---|---|---:|---:|---:|---:|---:|---:|
| 0 | warm | 10.2185 | 0.4605 | 3.3830 | 14.0880 | 0 | 41 |
| 0 | cold-verified | 10.2035 | 0.6085 | 3.3670 | 14.1740 | 8 | 43 |
| 1 | warm | 10.0710 | 0.5315 | 3.4115 | 14.0470 | 0 | 41 |
| 1 | cold-verified | 10.2045 | 0.6720 | 3.3900 | 14.2930 | 8 | 43 |

All 32 children have the ordered three-call publication sequence: 64 fsyncs
and 32 renames. Sixteen additional cold-source fsyncs and 168 other-PID events
remain explicit. This trace does not measure native durability cost or justify
removing synchronization. Raw and decoded evidence is retained in compressed
form; decompressed sizes and hashes are checked before temporary-file removal.

## Final custody and follow-up

The independent audit covers 76 successful build, workload and auxiliary
commands, 9,389 exact source inputs, all 34 reports and 444 unique measured
child PIDs. Nine strict report mutations and seventeen profile/parser/census
checks pass, as does independent owner count/period agreement. The validation
ledger retains 36 offline invocations: 25 successes and eleven corrected
failures, each linked to its successful replacement check. The full reader,
profile analysis, owner census, mutation checks and audit replay after cleanup.
The audit's initial final-preview failures exposed historical quality-command
paths counted as current commands and the known availability/build overlap;
both failures and corrected audit versions remain retained.

Cleanup removes only the two marker-bound roots: 3,583 files / 3,224,443,012
logical bytes. Raw perf/trace data remains compressed and byte/hash verified;
executable identities, Cargo fingerprints and bounded disassembly remain
retained as evidence. Two temporary parser-snapshot directories were copied
into the packet and removed separately. The error JSON produced at the
repository root by the first trace-reader invocation was also archived and
removed. Unrelated workspace edits are preserved.

Changed-Part Deflate work is the next experiment for this fixture. Any candidate
must retain exact no-op, typed-refusal, preservation, cancellation, resource and
default durability contracts and pass matched end-to-end measurement. This
single-host synthetic diagnosis does not close broader CRUD/corpus, concurrency,
scaling, physical-I/O or cross-platform goals.
