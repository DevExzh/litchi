# 0788: cached-Part memory attribution

**Diagnostic only; production restored.** This batch does not reproduce or
explain the small-case RSS increase that rejected change 0787. It establishes
an important measurement discrepancy: GNU time's reported maximum RSS is
below acknowledged point-in-time residency in every one of the 32 instrumented
children. The earlier rejection remains authoritative; no candidate is adopted
and no adoption threshold is changed.

## Scope and evidence

The paired source is `749c5e1945`. Both legs rebuild the exact production-file
maps from 0787: the baseline and the archived all-ready cache scheduling hint.
The candidate changes three OPC source/test files; all other production files
and the ordinary native benchmark are identical. After capture, all 9,196
production-file hashes are restored to baseline. The new phase probe exists
only in the evidence packet.

The host is the same 32-core AMD EPYC 9R45 Linux environment, using Rust 1.95.0,
glibc 2.43, 4 KiB pages, and CPU affinity 0–31. Inputs are deterministic,
immutable in-memory packages with 32 selected Parts. Small members are 4 KiB;
the large control uses 256 KiB members. Fresh means a first payload operation
after metadata setup, not physical cold I/O. Primed means one verified preload
outside the measured interval. The native clock covers the public ordered read;
corpus construction, setup, preload, verification, and teardown are excluded.

The frozen matrix comprises:

- 120 native children: ten cases, six paired blocks, 30 samples and three warmups.
- 20 qualification children: one sample per case and leg.
- 64 diagnostic children: four cases, two repeats, one-sample/no-warmup and
  30-sample/three-warmup protocols, with the same probe's handshake on and off.
- 16 heaptrack children: four cases, two repeats, 30 samples and three warmups.

That is 220 reports and 5,092 measured samples. Native, phase-observer, and
profiler populations remain separate. The same-width floor-65,536 control
cannot enter the candidate cache scheduling hint. Every sample verifies exact
payload bytes/order and finite CPU-task/permit accounting. The observer sees
64 logical source calls for fresh reads and zero for primed reads, with no
short reads or outstanding reads after return.

## Fresh native pairs

The original rejecting case is small/primed/floor-0/width-4. Its six current
GNU-time RSS observations are:

| Block | Before KiB | After KiB |
| ---: | ---: | ---: |
| 0 | 4336 | 4188 |
| 1 | 4268 | 4296 |
| 2 | 4336 | 4264 |
| 3 | 4204 | 4228 |
| 4 | 4340 | 4148 |
| 5 | 4404 | 4152 |

The median paired after/before ratio is **0.974631**, with bootstrap 95%
interval **[0.949270, 1.006135]**. The historical 0787 ratio was 1.056969 with
all six pairs adverse. These are separate experiments; their samples are not
pooled, and the new result does not erase the earlier failed guard.

Cache-hit latency remains much lower: paired p50 ratios are 0.033102, 0.023912,
and 0.009718 for small primed widths 4, 8, and 32, and 0.031718 for large primed
width 4. The same-width small serial-floor control has p50 ratio 1.008163.
These are bounded low-level read results, not whole-document CRUD speedups.
The [complete table](results/change-0788/paired.md) and CSV retain all ten cases,
tails, faults, block distributions, and uncertainty. Bootstrap uses 10,000
resamples, seed 788078, and sorted endpoint indexes 249/9749.

## Residency and counter disagreement

The probe emits a phase record, waits for an exact acknowledgement, and resumes
only after the parent records `/proc` smaps, smaps_rollup, maps, status, stat,
and task IDs. PID, starttime, executable, parent, and phase order are checked.
All 432 snapshots have one joined benchmark task. First and last measured
samples are observed; transient worker peaks between markers are not captured.

For the rejecting case under the 30-sample protocol, the final measured
operation shows these after-minus-before differences:

| Repeat | smaps RSS KiB | Anonymous KiB | Executable-mapping RSS KiB |
| ---: | ---: | ---: | ---: |
| 0 | 0 | -4 | +4 |
| 1 | +4 | 0 | +4 |

Within each of those four children, RSS does not increase from the post-preload
marker to the post-operation marker, for either the first or last sample.
This does not establish absence of a transient allocation or peak. It shows
that the earlier 222 KiB median increase is not present as retained residency
at these observed phases. The one-sample controls and the more variable
width-32 anonymous mappings remain in the full phase analysis; generic
anonymous mappings are not assumed to be heap, TLS, or worker stacks.

In **all 32 handshake-enabled children**, GNU time's `%M` is below both the
largest observed smaps_rollup RSS and status VmHWM. The smaps-minus-time gap
ranges from **156 to 1,716 KiB**. For example, repeat-zero baseline small/primed
width-4 reports `%M = 4052 KiB`, yet its final snapshot has smaps RSS and status
VmHWM/VmRSS of **5312 KiB**. Its paired candidate reports `%M = 4396 KiB` and
also reaches a retained **5312 KiB** snapshot. Per-mapping RSS sums equal the
rollup, and status agrees with snapshot RSS; PSS differences remain within the
per-mapping rounding allowance.

This is a counter disagreement requiring investigation, not a claim of negative
instrumentation overhead or proof that the historical rejection was spurious.
Linux's [proc documentation](https://www.kernel.org/doc/html/latest/filesystems/proc.html)
explains that scalable RSS accounting can be imprecise and distinguishes smaps
page-table scanning. That documentation motivates independent measurement; it
does not identify the cause of these particular observations. The handshake
also changes process scheduling and can fault pages, so its readings are
not substituted into the native population.

## Allocation and executable evidence

Heaptrack profiles include the whole child, including corpus generation,
metadata, preload, verification, and report serialization. The unchanged
corpus builder retains payloads plus both OPC and CFB containers even though
this matrix executes only the Parts route. Its unmerged peak
stack costs are intercepted live allocation bytes, not RSS or operation-only
allocation. Raw compressed traces and print commands are retained. Allocation
size histograms are derived offline from those same traces, without new native
runs.

For small/primed/floor-0/width-4, allocation calls fall from 134,931 to 132,666
in repeat zero and from 134,935 to 132,667 in repeat one. Histogram-weighted
allocated bytes fall from 109,228,464 to 108,937,894 and from 109,228,848
to 108,937,990 respectively. Nevertheless, summed
peak stack costs rise from 981,543 to 1,008,515 bytes and from 980,042 to
982,353 bytes. The peaks include cold preload/worker/decompression allocations;
they do not isolate the cache-hit helper. The serial-floor control retains
130,355 allocation calls on both legs and repeats. Rounded human-readable
heaptrack sizes are distinguished from exact stack-cost and histogram sums.

Both native and diagnostic ELF files grow 19,672 bytes, including debug and
other nonresident content. Allocated classified sections grow only 4,104 bytes:
text +2,864, rodata +88, bss +1,056, and eh_frame +96. This is consistent with
a small executable-residency contribution, but neither file size nor these
section totals establish the cause of the historical RSS delta.

## Validation, corrections, and architectural scope

Six fresh probe checks pass: format, all-feature/all-target compilation,
five tests, all-feature/all-target Clippy with warnings denied, rustdoc with
warnings denied, and the repository boundary checker (65 packages, 244
internal declarations, 11 explicit existing debt items). The exact historical
candidate's six production gates and 994 passed/one ignored tests are reused
through source and artifact hashes; they are not represented as new test runs.

The capture runner initially mislabeled GNU time's `%R` minor-fault count as
`elapsed_seconds` in ten baseline qualification JSON sidecars. Their raw
`%M %R %F` files and receipts remain unchanged. The original frozen driver and
an explicit correction receipt are retained; the label was fixed before any
remaining capture. Offline replay decodes the raw counters and confines that
legacy alias to those ten records. No elapsed-seconds measurement is claimed. Offline checker corrections for
rollup headers, heaptrack labels, restoration, and cleanup custody are recorded
in `replay-corrections.json`; raw measurements and policy remain unchanged.

All 35 previously reviewed architecture/goal input hashes remain unchanged.
The scoped ADR assessment is:

| ADRs | Assessment |
| --- | --- |
| 0001–0004, 0007, 0024–0027 | No retained public API, semantic model, transaction, or ownership change. |
| 0005, 0006, 0008 | Measurement scopes and perturbations are explicit; exact outputs, source fences, and resource checks remain required. |
| 0009–0023, 0028–0030 | No retained container, format, ODF, iWork, crypto, or publication change. |
| 0031 | Exact historical execution admission, task charges, worker/I/O budgets, and permit release are verified. |
| 0032 | No new derived cache or retained memo is introduced. |

This is CRUD taxonomy category 15 diagnostic evidence. Public Office CRUD,
physical-cold/range sources, cross-session contention, hardware-counter breadth,
and the program's larger goals remain open. OLE2/OOXML stay active, ODF remains
deferred by the recorded owner decision, and iWork is excluded.

## Next action

The original small-case memory cause remains unresolved. Before reconsidering
this candidate, qualify a native memory measurement that reconciles the parent
exit-time accounting with observed child residency, and distinguish code-layout,
allocator scheduling, and transient peaks using controlled measurements.
Any new adoption experiment needs a newly frozen representative policy; this
packet supplies neither adoption nor a weakened interpretation of 0787.

The [packet](results/change-0788/README.md) retains exact sources, raw captures,
allocation traces, phase snapshots, corrections, and offline replays.

## Integration and cleanup

All four executable identities were checked before removing 847,553,767 bytes
of owned build output. The ordinary, phase, and allocation replays and combined
validator pass after target removal. The packet keeps a cleanup witness for
executable custody, and production remains the qualified baseline.
