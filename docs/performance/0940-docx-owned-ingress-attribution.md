# 0940: attribute DOCX owned-ingress preparation and opening

The new diagnostic harness identifies two follow-up targets: slow reader
preparation samples coincide with hundreds of minor faults, and reader opening
performs 256 more reallocations than vector opening. This batch adds measurement
and validation tooling; it changes no production library behavior and qualifies
no optimization. The [0937 rejection](0937-docx-combined-fragment-validation.md)
and its original gates remain in force.

## Capture and scope

One authorized capture completed all 80 children, without retries or replacements:
72 workloads (three arms, twelve blocks, two routes) and eight route-independent
calibrations at boundaries 0, 4, 8 and 12. Every child discarded five warmups and
retained 50 samples. All six arm permutations occurred twice; route order
alternated by block. Children ran sequentially on CPU 12 with a 900-second deadline
and one-second observation interval. This is a fresh-process, warm-cache experiment
on an AMD EPYC 9R45 host, not a cold-storage or exclusive-host measurement.

The normal binary supplies both probes-off and probes-on arms. A second binary
adds allocator instrumentation. Probes-off performs no allocation-region or procfs
snapshot calls. Probes-on in the normal binary reports allocation counters as
unavailable, never as zero. Neither arm establishes heap/layout equivalence to the
historical 0938 binary.

The measured source is baseline `3ae856d76424f2f9b58d26dcd20fd074b8e97541`
plus the four harness files recorded in the retained source closure, **not a build
of this batch's parent HEAD**. Both builds used Rust 1.95.0, release, offline,
locked, `-j2`, with both Rust flag environment variables unset. The closure contains
10,780 files and 260,009,611 bytes. Build commands, environment and executable hashes
are retained in the authorized manifest and build receipts.

## Timing results

Values below are the median of twelve child medians, in microseconds. Operation
is the sum of preparation, open, query, edit, commit and publication time; wall
also includes observation overhead. Phase medians need not sum to operation medians.

| Arm | Route | Operation | Wall | Preparation | Open | Observer gap |
|---|---|---:|---:|---:|---:|---:|
| Normal off | Vector | 3053.130 | 3053.823 | 34.068 | 44.443 | 0.055 |
| Normal off | Reader | 3395.046 | 3395.756 | 35.245 | 378.822 | 0.060 |
| Normal on | Vector | 3071.623 | 3154.560 | 34.917 | 49.310 | 31.003 |
| Normal on | Reader | 3416.633 | 3503.258 | 35.285 | 384.339 | 31.078 |
| Allocator on | Vector | 3115.305 | 3200.086 | 35.658 | 55.455 | 31.365 |
| Allocator on | Reader | 3565.555 | 3652.539 | 136.710 | 393.202 | 31.370 |

Normal-on/off median within-block operation ratios are 1.006418 for vector
(descriptive bootstrap interval 1.004902–1.006866) and 1.008541 for reader
(0.944244–1.069545). Wall ratios are 1.033267 (1.031990–1.034158) and
1.033594 (0.967623–1.093945). The prospective reducer resamples twelve paired
blocks 10,000 times, using seeds 94010000 and 94010001. These intervals describe
this small, temporally related capture; they are not optimization acceptance tests.

Reader preparation remains variable. Counts at or above the prior 100 µs threshold
are 213/600 normal-off, 171/600 normal-on and 303/600 allocator-on; every vector
sample is below the threshold. Therefore aggregate reader medians can change
substantially as the proportion of slow samples changes.

## Allocation and fault attribution

These allocator counters are constant across all 600 samples per route and phase:

| Phase | Route | Gross requested bytes | Allocation calls | Reallocations |
|---|---|---:|---:|---:|
| Preparation | Vector | 6,369,855 | 2 | 0 |
| Preparation | Reader | 6,369,903 | 4 | 0 |
| Open | Vector | 285,401 | 388 | 39 |
| Open | Reader | 271,870,860 | 645 | 295 |

`allocated_bytes` adds the full new requested size on every reallocation, and
`allocation_calls` includes reallocation calls. These numbers do **not** measure
physical copies, memory traffic or RSS. The reader's extra 256 reallocations are
consistent with `read_limited` reserving exactly each additional 8 KiB read after
its initial reservation. Bounded buffer growth is a concrete next candidate,
subject to preserving size-limit and refusal behavior and fresh qualification.

An exploratory association splits preparation samples at 100 µs. The threshold
predates capture; selecting fault stratification was **post-capture**.

| Probes-on arm, reader | Samples | Median preparation (µs) | Median minor faults | Fault range |
|---|---:|---:|---:|---:|
| Normal, below threshold | 429 | 34.830 | 1 | 0–1 |
| Normal, at/above threshold | 171 | 242.981 | 482 | 481–483 |
| Allocator, below threshold | 297 | 35.050 | 1 | 1–1 |
| Allocator, at/above threshold | 303 | 241.721 | 482 | 482–483 |

Each probes-on vector arm has 600 below-threshold samples and one minor fault per
sample. Constant requested preparation bytes alongside this association motivate
further investigation; they do not establish an allocator or OS causal mechanism.
Pooled samples are not independent child replicates, and process intervals include
procfs observation work.

## Observation boundaries and correctness

The before-preparation process snapshot lies outside the allocation region. A
split ends preparation allocation accounting; a second split ends the observer
gap; `finish_split` ends open and combined allocation accounting. The gap includes
the after-preparation and before-open process snapshots. The after-open process
snapshot is outside combined allocation accounting but inside the combined process
observation. Scope strings record these different boundaries explicitly.

All phase, gate, allocation and process vectors share the operation sample order.
Wall samples have their own ranking. The reducer inverts both permutations before
combining observations. CPU ticks remain ticks; process high-water RSS is an
absolute observation; live-byte and peak values are not summed as additive work.

Calibration executes only black-box corpus observations, with no clone/open/edit
workload. Its operation median is 0.040 µs; normal wall medians at the four
boundaries are 49.431, 50.300, 49.105 and 50.205 µs; allocator wall medians are
51.175, 49.925, 51.190 and 50.900 µs. Calibration has a different heap context and
is reported separately, never subtracted from workload timings.

Each workload sample passes twelve semantic gate groups covering snapshot shape,
non-noop edit, forward/inverse/stale patches, inverse applicability, exact output,
reopening, exact source restoration, source immutability, cache contract and
determinism, sink bounds and reader activity. Correctness checks run after wall
timing; the cache observation between commit and publication remains in place.
Independent validation additionally enforces sink determinism and counter/status
contracts. Both packets passed the final independent checker: 80 children each,
identical child index, maximum report 518,324 bytes. Same-packet and wrong-terminal
negative controls reject substitutions.

Both normal and allocator focused Rust suites passed four tests each, and both
Clippy and rustdoc checks passed with warnings denied. All fourteen Python child
validation tests passed. The final report validator’s synthetic controls passed five positive and
eighteen rejection controls, recorded in external `validator-review-v3b.json`
(SHA-256 `60760b95e8cd30f4e0d0f5a862c3da9b6659d692f5c85d0cb9b12adb3e32e2ce`).
Actual Rust preflight used eight children with the
same five warmups and fifty samples; the runner's separate synthetic preflight is
not counted as Rust execution.

## Evidence and reproduction

[Reduced evidence](results/change-0940/reduced-evidence-v1.json) contains all group
statistics, paired ratios, calibration summaries, constant-counter checks,
per-child threshold counts, exploratory fault summaries, and artifact descriptors.
[Capture custody](results/change-0940/comparison-postcapture-custody-v2.json)
records both packet locations, replay, archive verification and cache disposal.
The verified archive is `/home/zhuhe/litchi-performance-0940-archive-v2.tar.zst`
(177,999,135 bytes; SHA-256
`3d9a975a7eb756448e14f4c7ff740df3ea72c61b8fa483f3ed2308a285dc10af`).
Its exact tar command is retained in `archive-scope-v2.json`. Verified cleanup
removed 6,326 intermediate files / 3,057,525,193 bytes while preserving both bound
release executables. External artifacts remain under
`/home/zhuhe/litchi-performance-0940`; no binary or evidence file of 1 MiB or more
is committed.

The authorized manifest is `manifests/formal-manifest-authorized-v2.json`
(38,997 bytes; SHA-256
`14c6c32810f03017711e2e4d86e30a984c35834d3eb649cba94224da92d250f4`).
The final full reduction is `descriptive-analysis-v2.json` (1,281,705 bytes;
SHA-256 `5c3b37cc4e4892a2660385788da638c5e233a8f7d87e285ca5fee706f1aa75d3`).
The source listing is `source-candidate-manifest-v1.jsonl` (1,665,554 bytes;
SHA-256 `8eb16940342230e36c4c121a71a49675e30b77b67a52f73e18b977fa1a100f1f`).
The listing is regenerated by retained `generate-source-candidate-list-v1.py`;
its exact source content is the retained closure. Build start receipts record
commands for regenerating the two executables; the final manifest binds them.

The commands below independently reproduce each reduction using retained canonical
inputs. The association and reduced-evidence commands read the retained v2
analysis, rather than the first command’s newly generated output:

```sh
cd /home/zhuhe/litchi-performance-0940/formal-packet-v1
python3 -B ../analyze-0940-v1.py ../analysis-input-index-v2.json ../descriptive-analysis-reproduced.json
python3 -B ../preparation-fault-association-v1.py ../descriptive-analysis-v2.json . ../preparation-fault-association-reproduced.json
python3 -B ../reduce-evidence-v1.py .. /tmp/litchi-0940-reduced-reproduced
```

The scripts are also committed beside the reduced evidence. The manifest and
outer-terminal receipts preserve the capture argv/environment and build commands;
re-executing workloads produces new measurements, not byte-identical timing data.
The full original and moved packets each contain 481 files / 45,653,539 bytes.
Their replay reduction is identical. Analysis v1 and v2 differ only in the bound
independent-review descriptor; all numerical fields are identical.

## Retained preparation history

Preparation failures and superseded receipts remain external. Seventeen early
unbound smoke files are excluded from accepted evidence. An accidental generated
source-tree target directory was removed, and the full source listing reverified;
proper checks used explicit external targets. Invalid rustdoc invocations were
retained, followed by successful normal v3 and allocator v4 checks. Early
preflight command/fixture failures were corrected before the sole formal capture.

The full source-listing receipt was written after build launch; its earlier
“before build” wording is corrected by the final manifest's timing note. The
source copy and four-file overlay preceded launch. Validator v3 was overwritten
by v3b during preparation: its old receipt survives, but the old v3 source bytes
are not claimed to survive. Final v3b bytes are manifest-bound. A schema-only
preflight review was superseded by the independently bound review. Checker v1's
path/outer-terminal controls were strengthened in v2 and the unchanged packets
reaudited. No failed preparation run was converted into a formal pass.
