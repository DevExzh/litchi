# 0843 — current-source many-Part OPC publication profile

This diagnostic identifies topology preparation as the largest observed
instruction region for a one-Part replacement in the 4,096-Part fixture:
59.18–59.34% of operation-scoped guest instructions. `PartNameIndex::insert`
accounts for 20.39–20.45%, within that region. The raw preservation index
accounts for 18.35–18.40%. No production change or speedup is claimed.

All 72 formal reports / 1,464 samples pass independent validation, with zero
planned native block-spread flags. Two observer compatibility preflights remain
separate. The measured source is commit `03fa423d73`; the broad performance
program remains active, with iWork excluded and ODF deferred by owner decision.

## Scope and protocol

The deterministic generic OPC packages contain 4, 64, 1,024 or 4,096 opaque
Parts, each 256 bytes, plus content types and package relationships. Names are
`/custom/p{index:05}.bin`; payload byte j is `(index * 17 + j * 13) % 251`.
Only the last Part changes, by XORing its first byte with 1. These are synthetic
low-level package operations in checklist category 15, not semantic Office
workflows or relationship-heavy graph edits.

The **publish** clock starts after source cloning, source-backed opening,
replacement-plan construction and empty output-Vec construction. It includes
`write_topology_to_stream`, topology validation/effective graph construction,
preservation-index construction, action planning, changed-member compression
and sequential output. **Lifecycle** additionally includes opening and plan
construction; its input clones and target URI are prepared before timing.
Hashing, output ZIP readback, exact oracle comparison and output dropping occur
after the clock. Consumed package state can be dropped inside the public call.
Do not subtract the two independently measured medians to estimate open time.

The expected output fixture uses a direct public writer call that bypasses the
non-inlined `profile_publish` owner. Assembly confirms one public writer call
inside that owner. Each output must match both the whole-archive and logical
identities. Independent ZIP readers reconstruct all payloads, verify CRCs and
metadata, and compare untouched raw local spans. The replaced target retains
stable metadata while its CRC/compressed size may change.

AMD EPYC 9R45, x86_64 Linux, rustc 1.95.0 / LLVM 22.1.2, release opt-level 3,
fat LTO, one codegen unit, default allocator, CPU 12. Probe registry versions
and checksums match the root lock. Six blocks alternate the order of eight
cases; each native process measures 30 samples after three warmups. Each
sample is followed by full validation, so these are warm process measurements
with validation-coupled cache and allocator state. The fixed matrix contains
48 native, eight qualification, eight Callgrind and eight Heaptrack processes.
All ran once, serially, without overlapping builds.

## Native results

Values are medians of six per-process nearest-rank quantiles. Intervals are
95% bootstrap intervals for the six p50 block values, using 10,000 draws and
seed 843083. All time values below are **milliseconds**. With only 30 samples
per process, p99 is the process maximum and is descriptive, not a universal
tail guarantee.

| Parts | Mode | p50 [95% interval] | p95 | p99 | Peak process RSS KiB |
|---:|---|---:|---:|---:|---:|
| 4 | publish | 0.025081 [0.024960, 0.025275] | 0.029990 | 0.039156 | 4,426 |
| 4 | lifecycle | 0.039390 [0.039020, 0.039940] | 0.050285 | 0.052890 | 4,370 |
| 64 | publish | 0.077121 [0.076781, 0.077730] | 0.087730 | 0.102506 | 4,488 |
| 64 | lifecycle | 0.180811 [0.180471, 0.181346] | 0.192561 | 0.199101 | 4,430 |
| 1,024 | publish | 0.886934 [0.885174, 0.888789] | 0.896114 | 0.901869 | 9,068 |
| 1,024 | lifecycle | 2.384535 [2.379305, 2.392235] | 2.402515 | 2.406020 | 9,114 |
| 4,096 | publish | 3.682305 [3.672660, 3.714830] | 3.752861 | 3.777676 | 22,460 |
| 4,096 | lifecycle | 9.813840 [9.797041, 9.855531] | 9.899676 | 9.939387 | 22,490 |

No max/min block-p50 ratio exceeds the frozen 1.20 review threshold. RSS and
page faults cover the entire child, including generator/oracle/validation;
they do not isolate publication memory. Every native process has zero major
faults. [All 56 metric rows](results/change-0843/tables.md) retain means,
fault counts and intervals. Independent raw-report arithmetic replay passes.
There is no before/after leg, regression comparison, geometric mean or
optimization-adoption decision in this batch.

## Instruction attribution

Callgrind collection toggles at the exact owner, with one positive owner call
and an empty terminal dump. Raw self-cost, direct-child and total conservation
pass for all eight captures. Ir is guest instructions, not native cycles or
elapsed phase time. Source-file labels can reflect inlining; the following
rows name concrete symbols. Inclusive regions overlap and must not be summed.

| Parts | Owner Ir, repeat 0 / 1 | Topology inclusive share | Preservation index inclusive share |
|---:|---:|---:|---:|
| 4 | 1,081,656 / 1,080,990 | 10.70–10.71% | 1.95–1.95% |
| 64 | 2,172,678 / 2,173,213 | 36.06–36.11% | 10.09–10.11% |
| 1,024 | 19,786,337 / 19,891,812 | 57.08–57.36% | 17.61–17.70% |
| 4,096 | 75,998,450 / 76,205,742 | 59.18–59.34% | 18.35–18.40% |

At 4,096 Parts, within topology preparation, Part-name insertion is
20.39–20.45% of total owner Ir, the two topology-limit scans are 11.54–11.57%,
and physical folded-name lookup is 7.77–7.79%. The low-level `write_to`
inclusive region is 11.65–11.89%, including plan preparation and output layout;
it is not a pure copy-time measurement. The current evidence ranks repeated
name/index construction ahead of emission. It does not prove those checks are
unnecessary. Tiny input has a different cost distribution: low-level
`prepare_with`, which includes generated-entry work, is about 72.6–72.7% Ir.

[Attribution](results/change-0843/attribution.json) retains every selected
symbol cost and both repeats. [Analysis](results/change-0843/analysis.json)
retains all self-cost rows, edges and file attribution. Differences between
repeats remain visible; no native phase-time conversion is performed.

## Allocation attribution

Heaptrack is a separate one-sample process population. Counts below include
allocation events whose full stack ancestry contains `profile_publish`.
Requested bytes include reallocations/growth and are not peak live bytes.
Both repeats agree exactly on these totals.

| Parts | Owner allocation events | Owner requested bytes | Topology events | Topology requested bytes |
|---:|---:|---:|---:|---:|
| 4 | 85 | 576,175 | 51 | 84,493 |
| 64 | 568 | 704,358 | 411 | 119,473 |
| 1,024 | 8,250 | 3,408,303 | 6,171 | 679,153 |
| 4,096 | 32,826 | 11,419,801 | 24,603 | 2,470,129 |

At 4,096 Parts the disjoint symbol-based groups also contain preservation
index 8,198 events / 950,728 bytes; plan 4 / 1,278,576; emission 15 /
6,359,663; and unclassified 6 / 360,705. These labels follow concrete stack
markers, not separately measured phases. The emission group includes output
buffer growth and generated-entry allocations. Its large requested-byte total
does not establish physical copy volume. Lifetime frees balance these matched
allocations but are not asserted to occur inside the timed operation.

## Decision, safeguards and next work

Retain this reproducible diagnostic without changing production. Investigate
bounded Part-name index construction first, then repeated metadata/limit
lookups and preservation-index ownership. A concrete small candidate is staging
only missing ancestor names: current `members.rs:201/209` allocates an ancestor
Vec and ancestor String for every insertion, then discards repeated ancestors
through `entry(...).or_insert(at)`. Both 4,096-Part Heaptrack repeats attribute
4,096 allocations / 98,304 bytes to the Vec and 4,096 / 32,768 to ancestor
Strings. All fixture Parts share `/custom`. The candidate must retain the first
descendant representative, duplicate/equivalent/derived conflict ordering and
fallible staging before index mutation. These are measured allocation sites,
not an estimate of removable native time. [Stack selection](results/change-0843/ancestor-allocation.json)
binds the exact source and both traces. Require a current-source controlled
candidate and end-to-end native/allocation evidence before adoption. Do not
replace the randomized hash strategy, skip validation, or introduce an
unbounded memo merely to reduce the measured instruction count.

| ADR boundary | Retained constraint |
|---|---|
| 0001 / 0002 / 0010 / 0011 / 0024 | Only packet-local probe/readers change; archive and OPC ownership remain intact. |
| 0003 / 0006 | Source identity, typed refusal, deterministic publication and untouched bytes remain required. |
| 0005 | Serial explicit input/output; no production allocator, hidden workers, ambient I/O or policy change. |
| 0008 | Synthetic substrate evidence is not native Office or complete CRUD coverage. |

[Source map](results/change-0843/source-map.md) records source freshness,
archive-end, folded-name ambiguity, signature/encryption, local/central span,
action coverage, resource and ZIP64 layout safeguards. Output layout must still
be proved before sink writes. Default durability is outside this in-memory
experiment and remains unchanged. No cold-cache, filesystem, remote-source,
concurrency, scaling, hardware-counter, peak-live-memory or copy-volume claim
follows.

Fresh probe formatting, warning-denied Clippy, five tests and release build
pass; the report/ZIP/Callgrind reader rejects 32 mutations and the Heaptrack
reader passes nine tests. Exact production-source identity permits reuse of
0842's 6,931 passing affected-owner tests and applicable gates, freshly replayed
before admission. Its baseline-confirmed iWork example/rustdoc exclusions
remain explicit. Two early probe Clippy failures and the superseded reader
qualification are retained, with no replaced workload. The final packet binds
all inputs, commands, failures, measurements, cleanup and exact commit paths.

[Packet and replay instructions](results/change-0843/README.md).
