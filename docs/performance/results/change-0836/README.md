# 0836 — OPC source-backed filesystem save profiling

This packet follows the admitted 0835 route/cache baseline. It profiles only
`opc_file_source_one_part_atomic_save`, with unchanged committed Rust source.
The completed matrix has 34 reports / 444 measured outputs: two qualification
reports / four outputs, 24 native control reports / 288 outputs, four CPU
profile reports / 120 outputs, and four syscall trace reports / 32 outputs.
Each population remains separate. iWork is excluded.

The baseline and frame-pointer builds use exactly the same 9,389 source inputs.
The second adds only `-C force-frame-pointers=yes`. Historical quality reuse
retains its original scope: the complete 569-pass/one-ignored suite preceded
two helper amendments, and the final helper passed twelve focused tests plus
format/check/Clippy/rustdoc/boundary gates. No new full-suite run is claimed.
Both fresh builds and exact output/cold-cache qualification are required.

Before capture, `symbols.py` admits emitted symbols through raw/demangled
address and size agreement, bounded disassembly, nested calls, and frame-pointer
prologues. The whole-save helper is inlined in the ordinary build, so the
existing single-Part overlay function is the fallback CPU scope. This excludes
package open and final atomic synchronization; separate ptrace syscall samples
observe file fsync, rename and directory fsync. Ptrace duration is not an
uninstrumented causal fraction and is never subtracted from native timing.

CPU profiles inherit across processes, then select exactly the measured child
PIDs and the admitted function frames in the retained executable. Parent,
priming, post-operation verification and other frames remain counted outside
the admitted population. Raw records, unknown/empty stacks, malformed/lost
records, sampled periods and leaf counts are conserved. CPU cycles sampling
cannot price blocked time. No production speedup or optimization is adopted.

Cold inputs retain the proven private EOCD-comment alignment padding from
0834/0835. Page-cache residency and process I/O do not prove physical-device
reads. Child peak RSS includes setup/cold-preparation history. These are
single-host synthetic-corpus diagnostics, not general Office certification.

## Retained result and offline replay

The [report](../../0836-opc-filesystem-save-profile.md) contains native controls,
conditional owner counts, syscall durations and claim limits. Deflate is the
largest observed overlay stack group. Unknown frames keep all CPU fractions
withheld. The full save helper is not the CPU scope.

`analysis-amendment.json` records offline corrections with immutable admission
snapshots. Runtime addresses use independent perf mmap/build-ID records and
ELF segment offsets. Cold traces retain one pre-operation source fsync per
child separately from the three successful atomic-publication calls. The
original plan and failed checks remain visible. No captured workload was
replaced. The sole command overlap is the perf-availability probe during the
frame-pointer build, before qualification; workload captures are serial.

From the repository root, these commands only read retained evidence:

```sh
python3 -B docs/performance/results/change-0836/replay_reports.py
python3 -B docs/performance/results/change-0836/profile_analysis.py --check
python3 -B docs/performance/results/change-0836/crosscheck.py --check
python3 -B docs/performance/results/change-0836/reader-tests.py --check
python3 -B docs/performance/results/change-0836/profile-tests.py --check
python3 -B docs/performance/results/change-0836/audit.py --check
python3 -B docs/performance/results/change-0836/seal.py verify --committed
```

The final command applies while this batch is HEAD. Source-aware readers
require the frozen source inputs to remain unchanged. Reproducing native
capture requires a fresh owned packet/target namespace; capture scripts refuse
to overwrite evidence and the original temporary roots have been removed.
`cleanup.json` retains deleted executable/fingerprint/raw-file descriptors and
compressed-data witnesses. The packet retains nine strict report mutation
checks, seventeen profile/parser/census checks, independent owner agreement,
and post-cleanup replay. The broader non-iWork performance goal remains open.
