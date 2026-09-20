# 0718 — DOCX serialization allocation and mapping attribution

Status: unchanged-source diagnostic evidence. No production optimization or
native speedup claim; 0715 remains rejected.

## Question and fixed protocol

The previous turn was progress. Batch 0717 found a recurring 78-minor-fault
regime in generated serialization, including mixed processes whose zero-fault
and 78-fault samples had different elapsed medians. It did not identify an
allocation owner or faulting instruction. This batch adds stack evidence using
the existing procfs-enabled harness, without changing its source.

A new release executable is built from revision `d48a92dc9a`, with exactly the
0717 source census and `ordinary-save-process-metrics` feature. The standalone
harness has its own workspace; root workspace release profile settings are not
assumed to apply. Host and tools are recorded: AMD EPYC 9R45, Rust 1.95.0,
strace 6.19 with libunwind, and Valgrind 3.26/DHAT. CPU 12 is pinned.

Two repeats each cover generated-medium and the 55,519-byte NumberedList
fixture, reversing corpus order in repeat two. Four mapping children use
100 warmups and 200 measured counting-publication samples each. Four DHAT
children use zero warmups and one measured sample each. The mapping lane finishes
before DHAT starts. All eight children and all raw artifacts are retained;
there are no replacement runs or adaptive profiling lanes.

Strace records `mmap`, `mremap`, `munmap`, `brk`, `madvise`, `mprotect` and
`openat` with up to 40 stack frames. Existing procfs opens form ordered
`io`/`stat`/`status` snapshot triplets. After 32 empty control pairs, each iteration
has a before and after triplet. The interval between the before-status open and
after-io open conservatively contains serialization plus the before-probe tail.
It is not the exact timer interval. The final 200 iteration windows align with
raw acquisition-order deltas; the first 100 remain warmup evidence.

An attributed serialization stack must contain both `write_plain` and
`ordinary_save::run_case`. The interval check further separates measured,
warmup, control and outside-window memory calls. Other publication stacks,
probe activity, unknown frames and truncated contexts remain separate. Missing
stack evidence cannot establish that an owner has no allocations or syscalls.

DHAT uses a replacement allocator and has no phase toggle. Only cumulative
allocation observations whose resolved stack contains both owner and caller
are selected. Whole-child totals remain distinct. These are not native mapping,
page-fault, RSS, operation-peak or latency measurements, and do not establish
complete coverage of all phase allocations.

## Symbol resolution

Strace left local Rust names blank but retained mapped-file offsets. All 188
unique local offsets are resolved offline against the exact captured executable;
no trace was rerun. The executable ELF load segment maps file offset `0xb92040`
to virtual address `0xb93040`, requiring a 4,096-byte adjustment before symbol
lookup. Feeding the raw offset directly to `addr2line` identifies unrelated
neighboring functions.

The mapping follows strace's recorded file-offset semantics, confirmed in its
[6.19 libunwind implementation](https://github.com/strace/strace/blob/v6.19/src/unwind-libunwind.c).
Each resolved name must also agree exactly with a containing `nm` symbol range.
Program headers, symbol table, input addresses, lookup output, tool identities,
source reference and binary-bound frame map are retained. Raw traces remain
unchanged, and replay requires no executable after cleanup.

## Results

All four traces validate 664 snapshot triplets: 64 control snapshots and 600
iteration snapshots. Each retains one extra status read outside the protocol.
All marker stacks resolve; none is truncated. Each trace also retains 27 other
unknown-context events, which cannot establish owner absence. All eight
published outputs pass normalized parity.

| Corpus / repeat | Zero-fault samples | Exactly 78 faults | Other fault counts | Phase memory calls in 200 measured windows |
|---|---:|---:|---|---|
| generated R1 | 0 | 195 | 5 × 79 | 200 `brk` |
| generated R2 | 1 | 189 | 9 × 79, 1 × 13 | 199 `brk` |
| NumberedList R1 | 197 | 0 | 1 × 1, 1 × 2, 1 × 8 | 0 |
| NumberedList R2 | 193 | 0 | 5 × 1, 1 × 2, 1 × 8 | 0 |

Every measured generated `brk` has a resolved `zlib_rs::deflate::init` stack
under serialization. R2's single zero-fault sample is acquisition index 0 and
has no phase memory call; index 1 has 13 faults and one `brk`. One zero-fault
observation is not a replicated causal comparison. No measured phase `mmap`,
`mremap`, `munmap`, `madvise` or `mprotect` occurs in these traces. Warmups
additionally contain 11 and 4 phase `brk` calls for generated R1/R2 and two
each for NumberedList; these are excluded from the measured counts above.

`brk-analysis.json` validates requested versus returned program breaks and
reconstructs transitions within each process. Every retained request succeeds.
Generated R1's 200 measured growth calls comprise 173 increases of 376,832 bytes
and 27 of 380,928 bytes. R2 has 198 increases of 380,928 bytes and one of
229,376 bytes. These are address-space break changes, not resident-byte or fault
counts, and are not summed into a claimed operation memory footprint.

Outside the conservative windows, 209 and 201 shrink calls respectively have
`systrim`/free stacks under `run_case`. Of those, 206 and 199 also show explicit
mutable-document/package destruction frames; the remaining 3 and 2 do not
identify that destructor. These totals include warmup and interstitial work.
The code drops the owner after the after-probe. This shows heap-lifetime changes
outside the timer alongside repeated construction growth inside it, without
proving that a different allocator policy would improve actual workloads.

DHAT repeats agree exactly on these cumulative allocation observations:

| Corpus | Whole-child bytes / blocks | Selected serialization-stack bytes / blocks | Deflate-init bytes / blocks |
|---|---:|---:|---:|
| generated | 19,162,168 / 86,842 | 801,740 / 6,522 | 380,032 / 1 |
| NumberedList | 14,799,218 / 25,103 | 1,113,016 / 752 | 760,064 / 2 |

The generated selection has 109 allocation contexts; NumberedList has 128.
Each Deflate initialization requests 380,032 bytes in DHAT through the same
construction chain identified by the trace. Whole-child and selected-context
totals remain separate: setup, opening, edits, probe reads and verification
also allocate. Fourteen unknown-context records per DHAT child remain visible;
none reaches the 40-frame limit. Peak fields are excluded from additive
summaries. These totals do not establish native allocation latency.

## Mechanism and decision

The source chain is DOCX `write_plain` → OPC `write_to_stream`/`write_counted`
→ preservation `prepare_with`/`generated_entry` → `Compress::new` →
`zlib_rs::deflate::init`. The generated workload regenerates one document member;
NumberedList regenerates two members. Unchanged members already use raw-copy
preservation, and exact-source passthrough cannot apply to this intentionally
edited workload.

The allocation owner is now evidenced rather than inferred from page counts:
fresh compressor construction requests 380,032 bytes per regenerated member,
and the traced generated regime repeatedly grows the program break along that
call path. The fixed matrix has only one zero-fault generated measured sample,
so it does not provide a strong within-trace comparison or explain 0716's slow
NumberedList process. No further profiling lane is added to seek a desired
regime. The useful follow-up is to account for untimed owner destruction and
process replication in future performance decisions, then return to ranked
work-elimination opportunities; this evidence alone warrants no production
allocator or lifetime change.

Cross-member Deflate pooling is not a new candidate. Batch
[0618](0618-zip-writer-deflate-state-reuse.md) retained authored-writer state
reuse and output-buffer work elimination but rejected preservation pooling after
its measured regression. The member-local state remains intentional. Likewise,
that batch measured the mini-archive parsing cost as small and retained its
structural validation. This packet does not relax either decision or infer a
profitable allocator policy from a fault association.

## Verification and limits

Every receipt passes source, executable, fixture, command, environment and
artifact checks. Published bytes, decoded target/member identities and edit
outcomes match the retained baseline. Eleven in-memory corruptions are refused,
covering statistics, CPU command, decoded sizes, instrumentation, counter
alignment, controls, DHAT counters/frame indices, ELF address conversion,
marker order and failed `brk` requests. Both derived analyses replay exactly.

Production and Rust harness source are byte-identical to 0717, so their prior
verification is explicitly reused. The 0717 harness results are 531 default tests,
530 feature tests (one opt-in test ignored in each), 89 latency-guard tests,
warning-denied Clippy/rustdoc and six repository gates. The nested 0713 production
record covers 4,995 tests and 92 passed/46 ignored doctests, with seven baseline-proven
PPTX/XLSB test Clippy exceptions disclosed. No reused suite is represented as
newly run. The diagnostic build, eight profiler runs, parser checks, final
report gate, replay and cleanup audit are fresh.

All three owned scratch roots are removed with the executable identity retained
as a cleanup witness. The final artifact seal includes the raw stack evidence,
offline resolution witnesses, scripts, derived analyses and review.

The traced windows include probe tail work and tracing perturbation. A mapping
request is not a physical-page count, and a syscall stack is not the instruction
that faulted. DHAT changes the allocator. No comparison between these lanes
establishes a speedup, no counter is divided by native elapsed, and no generic
memory or cold-cache conclusion follows. Rust hash seeds, allocator/ASLR layout
and unlisted environment fields remain uncontrolled. Prior 0716/0717 process
variation and 0715's rejection remain intact. The non-iWork goal remains active.

[Evidence and replay instructions](results/change-0718/README.md).
