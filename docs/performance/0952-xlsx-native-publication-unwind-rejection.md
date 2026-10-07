# 0952 — rejected native XLSX publication attribution

The first native profile fails its frozen evidence-quality gate: **zero
publication-owner samples**, against a required minimum of 100. All 200
semantic samples pass, but their callchains cannot establish a native hotspot
ranking. The remaining three planned captures were not authorized or run.
There is no retry, replacement capture, production change, adoption, latency,
allocation, or speedup claim. iWork remains excluded.

## Question and source review

0951 ranks preservation writing at 57.70–60.95% of publication guest
instructions, with Deflate longest-match search the largest exclusive row.
That does not establish a native-time fraction, particularly with runtime CPU
dispatch enabled. This batch attempted to observe native user-cycle stacks
before changing the compression path.

The current path is XLSX `publish_multi_commit_to_stream`, OPC topology
preparation and preservation writing, then soapberry's `generated_entry`.
Changed entries retain the source compression method and use shared logical
payloads; untouched entries use `Copy`. Each regenerated Deflate member has
one compression pass. Source inspection does not identify repeated compression
of the same member. The measured constructor is unmanaged and attaches no
write session, so it cannot use the existing opt-in managed compression wave.

Prior rejections still apply: 0618 rejected carrying a compressor across a
preservation plan after native regressions, despite fewer guest instructions;
0752 disproved byte transparency of changing compressor input boundaries;
0837 found that fresh-only level-5 publication broke serialized inverse
restoration. The current evidence does not authorize those shortcuts.

## Frozen method and custody

The batch started at `2d6833be51615486ef3c2a83d4791fd2a3f223c4`. Its delta from the
0950 adopted source contains only 13 performance documentation/evidence paths.
The retained 0950 executable is 266,090,728 bytes, SHA-256
`198c2e87471a088cf054a23785e4068eaf7f205ffb1fb25cad5ba0b7a6db5bec`.
It was built from the frozen 0950 candidate, not at current HEAD. All 10,790
retained source files were rehashed; the previously declared ignored root lock
exception remains. The standalone tracked lock controls the `--locked` build.

The prospective schedule was medium, dense-sparse, dense-sparse, medium, each
with 200 samples and zero warmups. Sampling used CPU 12, `cycles:u` at 997 Hz,
DWARF stacks of 32,768 bytes, no inheritance, a 900-second process-group
deadline, and fixed C/UTC environment. Record and decoder terminals precede
report consumption. The first child required semantic and raw admission before
any remaining child. The selected owner excludes setup, open/planning,
lifecycle gates, reopen checks, and returned-snapshot destruction. The native
publication timer includes that destruction and is a different boundary.

The executable has symbols, debug data, and CFI covering all four owner
instances and the measured parent. Its owner prologue does not establish a
frame pointer, so the plan chose DWARF. Section presence and covering CFI proved
insufficient for reliable callchain reconstruction in this local toolchain.

## Retained decode failures and rejected population

The original `perf script --debug-mode` returned success and reported zero
misordered timestamps, but produced an empty sample file: that option performs
timestamp checking rather than printing samples. A separately terminaled
offline export without that option decoded the **same** `perf.data`; the empty
original and its receipt remain. A separate mmap/task export retains the
recorded address mappings. Neither operation reran the workload. The original command and all derived
exports retain separate terminal receipts.

| First medium child | Observation |
| --- | ---: |
| Semantic samples | 200, pass |
| Sample headers | 8,879 |
| Sum of sampled period weights | 34,521,187,678 |
| Physical frames printed | 11,958 |
| Empty callchains | 1,682 |
| Exact mapped publication-owner frames | 0 |
| Admitted publication-owner samples / period | 0 / 0 |
| Required owner samples / period | 100 / 100,000 |

All header period weights are retained. These are sampled event weights, not
an exact hardware-cycle count or a wall-time measurement. Missing owner frames
are unknown membership, not evidence that publication did not execute. No
whole-process leaf ranking is substituted for the missing owner population.

Independent ELF/mmap inspection identifies a local one-page translation
mismatch. The executable RX map starts at `0x5b8f9e1f4000`, file offset
`0x8a6000`. Its ELF RX `PT_LOAD` has offset `0x8a6df0` and virtual address
`0x8a7df0`, giving page bases `0x8a6000` and `0x8a7000` and a correct load bias
of `0x5b8f9d94d000`. Runtime sample `0x5b8fa085f058` therefore maps to
`0x2f12058`, `longest_match+0x1a8`. The verbose unwind stream instead reports
`0x2f11058`, one page lower, in the preceding `deflate_slow` function.

The correct address has CFI in `0x2f11eb0..0x2f12281` with CFA `rsp+160`;
the wrong address selects the preceding FDE and different stack rules. The
verbose decode includes invalid/out-of-range memory accesses and zero return
records. This is a bound local diagnosis, not a claim about every perf build
or every missing frame. A draft offline mapper also encountered two ELF
segments sharing a file-offset page; executable-map permissions must select
the RX segment before calculating its bias. Unresolved ambiguity must reject.

## Decision and next experiment

Retain the failed first capture and stop this frozen schedule. Native
publication ranking remains unproven; 0951 is still instruction evidence only.
Next build a separate profiling executable from the same frozen source with
`-C force-frame-pointers=yes`, verify owner/parent frame-pointer prologues and
ELF mapping bindings, and prospectively qualify its first owner population.
Its instrumented timing must not become the ordinary native baseline.

After reliable attribution, census the changed Deflate task sizes before
qualifying the existing managed write wave. Require task count, total/largest/
remainder bytes, admitted width, budgets, output equality, and native/allocation
results. Preserve source checks, complete original auditing, compression and
container contracts, exact no-ops, reversible patches, bounded resources,
canonical fallback, and both one-edit controls including the 0950 tail flag.

## Reproduction and evidence

Raw capture, corrected exports, verbose unwind diagnostics, and failed reader
attempts remain external under `/home/zhuhe/litchi-performance-0952`.
Archive identity, final reader rejection, relocated replay, corruption checks,
and owned-scratch cleanup are recorded in the final accompanying closure.

The intermediate helper revisions behind early reader error outputs
were not retained separately. Those error outputs are historical diagnostics,
not independently replayable helper revisions. The frozen final v2 and derived
v3 readers are retained by hash; final replay covers the final rejection.
The v2 custody-limit receipt and retained v3 mapping error output make this
distinction explicit; reproducibility is claimed for the frozen final readers.

Final reader v3 (`321e415fd3092b3002c76114e7cb147f7382d231c697dbd2878d7bc9c9ddc08e`)
reproduces all 8,879 samples and the rejected zero-owner population. Its
19,038-byte result has SHA-256
`24c6f18f3c4d48278d877dc74a0ac200e8c798da7131eb4e051f7fbf926defdc`.
A synthetic owner/parent pair under the actual executable mapping qualifies;
wrong parent DSO, out-of-range IP, and wrong PID do not. Wrong build ID and
wrong event are rejected by the actual parsers. Ambiguous or non-executable
maps cannot qualify owner frames. The synthetic one-sample positive remains
below the capture admission minimum and does not admit a real profile.


### Archive closure

The external archive `archive-0952-v1.tar.zst` is 10,936,839 bytes, SHA-256
`e6fe78b8fa765a3c5247411641cca881f1c3fbf1a85666aac717413cf77f90db`,
under `/home/zhuhe/litchi-performance-0952-archive`. Its
[manifest](results/change-0952/archive-manifest.json) identifies all 68 files
(312,086,172 logical bytes), including every large raw artifact by path, size,
and SHA-256. No binary or large raw evidence is committed.

Extraction verified the complete file set, hashes, sizes, modes, and absence
of symlinks before reader execution. Relocated replay returned the expected
code 2 and reproduced the final 19,038-byte rejection byte for byte. The
archived `00b` and final `00c` results have the same hash. File-access tracing
recorded zero accesses beneath the original 0952 packet. Eight actual manifest
corruption controls rejected missing/extra files, file/directory symlinks,
wrong hash/mode, duplicate entries, and path escape.

The frozen reader has hardcoded paths; its `--packet` option does not relocate
it. The separately retained replay adapter remaps packet-local reads and
normalizes descriptor paths while preserving the declared external 0942,
0950, and 0951 dependencies. This is adapter-assisted replay, not a claim that
the original reader is portable. The reader controls and archival controls
verify different boundaries.

See the [closure](results/change-0952/archive-closure.json) and
[replay receipt](results/change-0952/replay-receipt.json) for exact commands,
helper identities, and limitations. To recover every raw artifact exactly,
extract the archive into a fresh directory with `tar --extract
--same-permissions --no-same-owner --file
/home/zhuhe/litchi-performance-0952-archive/archive-0952-v1.tar.zst
--directory <fresh-directory>` and verify against the manifest before parsing.
The capture runner and terminal receipts inside the archive retain the
original generation commands; this closure does not authorize another capture.
Repacking requires a fresh archive output directory and is not promised to
produce the same compressed bytes.

Owned replay scratch and bytecode were removed after verification; canonical
raw evidence and the compressed archive remain. Production source is unchanged,
so validation here consists of semantic, reader, manifest, and replay checks.
