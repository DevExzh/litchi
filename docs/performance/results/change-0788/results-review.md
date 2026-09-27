# Independent raw review: change 0788

This review covers the evidence packet for the cached-Part attribution
experiment. The current worktree history is at `749c5e1945` (the 0787
experiment integration and cleanup record); the archived candidate under test
is the rejected 0787 candidate from `f50e22fc3`. The candidate is attribution
evidence only. Production was restored after capture, and this review did not
build, capture, profile, or alter the production sources.

The raw evidence passes custody, payload, cardinality, and finite-resource
checks. Its diagnostic observations do not identify the cause of the old RSS
result. The 0787 rejection remains authoritative: this packet does not adopt
the candidate or change an adoption threshold.

## Independent replay of the packet

The plan has four deliberately separate lanes:

| Lane | Reports | Samples per report | Total samples | Design |
|---|---:|---:|---:|---|
| Native | 120 | 30 | 3,600 | 10 cases, 6 paired blocks, 3 warmups, both legs |
| Qualification | 20 | 1 | 20 | 10 cases, one before and one after |
| Memory | 64 | 1 or 30 | 992 | 4 cases, 2 repeats, off/on, two protocols, both legs |
| Heaptrack | 16 | 30 | 480 | 4 cases, 2 repeats, paired legs, 3 warmups |
| **Total** | **220** |  | **5,092** | 432 on-mode phase snapshots |

I walked each receipt-declared artifact reference independently against its
packet-local path or recorded cleanup binary witness. Every packet artifact was
present and every recorded SHA-256 matched; there were no missing packet
artifacts or recorded SHA mismatches. Receipt manifests repeat shared artifacts
across children and lanes, so this reference-level check is distinct from the
de-duplicated custody census. The reports retain the expected execution schema
and the exact sample counts above. Native and heaptrack legs use the recorded
before/after binaries; memory legs use the separately recorded diagnostic
binaries. The source-observer and phase-probe populations were not pooled with
the native population.

I reconstructed the deterministic payload oracle from the probe source rather
than trusting only the report flags. It uses the 32 member indexes, the
4-KiB/256-KiB member sizes, the member labels and offsets, and the little-endian
index-plus-payload sequence hash. The resulting sequence hashes are:

| Corpus | Logical bytes | Sequence SHA-256 |
|---|---:|---|
| Small | 131,072 | `5db97cf696709a91ead89194564cbc26a1ec5fa2309aa0043e97a3f3f38d02ed` |
| Large | 8,388,608 | `a238675779a8f3525cc743bcf06d07d2db3742b1c2b9c7bc962aa49b4b4f6a11` |

The mixed sequence is an oracle-only reconstruction; no mixed case was captured
in the 0788 matrix. All 220 reports and all 5,092 samples matched the
independent member SHA list,
member byte sizes, selected count, logical byte count, ordered flag, and
sequence hash. No payload or ordering failure was found.

The finite resource fields also replay. Feature-off native, qualification, and
heaptrack reports correctly expose no source metrics. Feature-on memory reports
show 64 logical source calls and 73,590 requested and returned bytes for fresh
samples, with zero short reads, zero active reads at return, and maximum
simultaneous reads of one to three. All primed samples have zero source calls,
bytes, short reads, active reads, and maximum simultaneous reads. Fresh reports
move from zero to 32 CPU tasks across the operation; primed reports move from
32 to 64. Every report records worker/I/O release and CPU-task-limit success.
The recorded limits include one million CPU tasks and a 16 MiB
`max_in_flight_bytes` limit. The latter is a scheduler limit, not an observed
total allocation or a process-memory cap.

## Native paired results

For each case and each of the six blocks, I computed nearest-rank p50, p95,
p99, then paired the after value with the before value. The table reports the
median of the six after/before ratios. Each bracket is the 95% bootstrap
interval from 10,000 resamples of those six paired block values, using seed
`788078` and sorted endpoint indexes `249` and `9749`. Latency values are
ratios of nanosecond quantiles; RSS is a ratio of the GNU-time maximum RSS
field. The raw six-block values and deltas remain in `analysis.json` and
`paired.csv`.

| Case | p50 ratio [95% interval] | p95 ratio [95% interval] | p99 ratio [95% interval] | RSS ratio [95% interval] |
|---|---|---|---|---|
| small fresh floor 0 width 1 | 1.002650230 [0.998056777, 1.005782763] | 1.004253023 [0.996071302, 1.019354174] | 0.990093897 [0.917331981, 1.006914836] | 0.997270246 [0.964559925, 1.008557852] |
| small fresh floor 0 width 4 | 1.005128095 [0.979689705, 1.015379601] | 0.972569533 [0.938838059, 1.039947778] | 1.019775709 [0.915635787, 1.063040675] | 0.971715926 [0.951148464, 1.001915220] |
| small fresh floor 0 width 8 | 0.996990598 [0.956005503, 1.049569855] | 0.970024443 [0.942219283, 1.044057513] | 0.950055299 [0.885634819, 1.033754483] | 1.003880099 [0.983843311, 1.020970673] |
| small fresh floor 0 width 32 | 0.996800583 [0.944531389, 1.011676246] | 0.967332441 [0.738890414, 1.046393655] | 0.959495894 [0.683121775, 1.035155297] | 1.059400673 [0.949460785, 1.099687747] |
| small primed floor 0 width 1 | 1.008113624 [0.997959184, 1.018527300] | 0.999977200 [0.982819870, 1.041957027] | 0.996138996 [0.619405594, 1.039309429] | 0.980180077 [0.952502223, 1.024736083] |
| small primed floor 0 width 4 | 0.033101616 [0.029085946, 0.036850873] | 0.033232126 [0.029857681, 0.038944521] | 0.040568887 [0.032435340, 0.081766138] | 0.974630996 [0.949269830, 1.006134649] |
| small primed floor 0 width 8 | 0.023911887 [0.022352790, 0.024775372] | 0.025979736 [0.025746306, 0.027779724] | 0.026881781 [0.024989000, 0.028505296] | 0.987291632 [0.952505342, 1.029600398] |
| small primed floor 0 width 32 | 0.009718446 [0.008991580, 0.010018277] | 0.009861446 [0.008090049, 0.012478790] | 0.010029087 [0.006485250, 0.016143121] | 0.965782888 [0.894198719, 1.049268681] |
| large primed floor 0 width 4 | 0.031717502 [0.030471908, 0.032510612] | 0.032196418 [0.030789307, 0.036278276] | 0.032061288 [0.030231576, 0.038146003] | 0.998059478 [0.993264598, 1.003018947] |
| small primed floor 65,536 width 4 | 1.008163401 [1.002032520, 1.014294282] | 1.019516772 [0.980999665, 1.045006127] | 1.018997224 [0.984725513, 1.038827557] | 0.968441798 [0.942716621, 1.017614033] |

The cache-hit p50 ratios are approximately 0.0331, 0.0239, and 0.0097 for
small primed widths 4, 8, and 32, and 0.0317 for the large width-4 case. The
same-width floor-65,536 control has p50 ratio 1.0082, so the cache-hit effect
does not appear in that serial control. The largest point estimates in the
table are p50 1.0082, p95 1.0195, p99 1.0198, and RSS 1.0594. Every RSS
interval includes 1; the cache-hit latency intervals are intentionally far
below 1. Fresh latency intervals are near one. The serial control's p50 is
1.0082 with interval [1.0020, 1.0143], just above 1; its p95 and p99 intervals
include 1. Major faults are zero in every native block, so no major-fault ratio
is defined; minor-fault raw values and paired deltas are retained in the
packet.

The current small/primed/floor-0/width-4 RSS block pairs are:

| Block | Before KiB | After KiB |
|---:|---:|---:|
| 0 | 4,336 | 4,188 |
| 1 | 4,268 | 4,296 |
| 2 | 4,336 | 4,264 |
| 3 | 4,204 | 4,228 |
| 4 | 4,340 | 4,148 |
| 5 | 4,404 | 4,152 |

Their paired median ratio is `0.974630996`, with interval
`[0.949269830, 1.006134649]`; the paired median RSS delta is -110 KiB, with
bootstrap difference interval [-222, 26] KiB. This is a separate current
experiment. It does not pool with, erase, or weaken the historical 0787
result.

## Phase snapshots and accounting

The enabled diagnostic lane contains 32 children and 432 snapshots. The
disabled lane has no phase snapshot series by design. Every enabled child has a
stable PID/starttime, executable, and parent identity; its task list contains
only the benchmark PID, and `status`, `stat`, and task identity agree. The
phase transcript and snapshot phase sequence agree:

`startup`, `corpus_ready`, `warmup_done`, `package_ready`, `after_preload`,
`after_operation`, `after_batch_drop`, `after_package_drop`, `samples_done`,
`report_written`, `report_dropped`.

The per-mapping RSS sums equal `smaps_rollup` in all 432 snapshots. `status`
VmRSS agrees with the rollup, while PSS differences stay within expected
per-mapping rounding. These are point-in-time observations: the handshake does
not record an unobserved allocation peak between markers, and two repeats do
not form a confidence interval.

For the 24 after/before pairs at each phase (48 endpoint observations across
both legs), the phase RSS medians are near zero but the ranges are variable.
The package-ready median is +6 KiB (range -68 to +328); after-preload is +2
KiB (range -324 to +204); after-operation and both drop phases are each +2
KiB (range -220 to +308). These ranges are comparable to the old 222-KiB
result and do not support a causal attribution. At after-operation, the
executable mapping has median +4 KiB (range -60 to +88), unnamed anonymous
mapping has median 0 (range -196 to +308), and heap, shared-library, and stack
categories have median 0 with small single-digit-KiB ranges. The mapping
categories do not establish ownership or causality.

For the historically rejecting case under the 30-sample protocol, the final
operation endpoint differences are:

| Repeat | RSS KiB | Anonymous KiB | File-backed KiB | Executable mapping KiB |
|---:|---:|---:|---:|---:|
| 0 | 0 | -4 | +4 | +4 |
| 1 | +4 | 0 | +4 | +4 |

Within these four children, post-preload to post-operation RSS is unchanged at
the observed endpoint for both first and last measured samples. That does not
rule out a transient peak, and generic anonymous bytes are not relabeled as
heap, TLS, or worker stacks.

The strongest diagnostic finding is a counter disagreement. In every one of
the 32 enabled children, GNU time `%M` is below both the largest acknowledged
`smaps_rollup` RSS and the largest `status` VmHWM. The gap ranges from 156 to
1,716 KiB. Two exact examples are:

| Child | GNU time `%M` KiB | Largest smaps/status value KiB |
|---|---:|---:|
| Repeat 0, baseline, small primed width 4, 30 samples | 4,052 | 5,312 |
| Repeat 0, candidate, small primed width 4, 30 samples | 4,396 | 5,312 |
| Repeat 1, baseline, small primed width 4, 30 samples | 4,424 | 5,304 |
| Repeat 1, candidate, small primed width 4, 30 samples | 4,492 | 5,308 |

The smaps/status maximums agree within each row. This disagreement needs a
controlled accounting investigation; it is not evidence that the observer
measurement explains or invalidates the 0787 rejection. The handshake itself
perturbs scheduling and page faults, so these snapshots cannot replace the
native process population.

## Heaptrack and executable checks

All 16 heaptrack children completed with exit and print status zero. Every
print command used `-m 0`, so merge-backtraces was disabled; the retained
profiles use peak-cost stacks. The profile covers the whole child, including
corpus construction, metadata, preload, verification, operation, and report
serialization. Source metrics are unavailable in this native feature-off
lane, so there is no source observer mixed into the profile.

The exact unmerged peak stack-cost pairs (after minus before) are:

| Case | Repeat 0 delta | Repeat 1 delta |
|---|---:|---:|
| small fresh floor 0 width 4 | -332 B | +1,902 B |
| small primed floor 0 width 4 | +26,972 B | +2,311 B |
| small primed floor 0 width 32 | +33,870 B | -29,718 B |
| small primed floor 65,536 width 4 | -1 B | -1 B |

For the rejecting small/primed/width-4 case, allocation calls decrease from
134,931 to 132,666 in repeat 0 and from 134,935 to 132,667 in repeat 1.
The offline histogram totals likewise decrease from 109,228,464 to
108,937,894 bytes and from 109,228,848 to 108,937,990 bytes. The peak stack
cost nevertheless increases from 981,543 to 1,008,515 bytes in repeat 0 and
from 980,042 to 982,353 bytes in repeat 1. These are whole-child intercepted
allocation observations, not RSS and not operation-only attribution; they do
not explain the earlier 222 KiB whole-process RSS delta. The serial floor
control has 130,355 allocation calls on both legs in both repeats and a
one-byte peak-stack decrease.

The native and diagnostic ELF files each grow 19,672 bytes in total file size.
Allocated sections grow by only 4,104 bytes: text +2,864, rodata +88, bss
+1,056, and eh_frame +96, with data unchanged. This is compatible with a
possible small code-layout contribution, but file size and section totals do
not establish the RSS cause.

## Correction and restoration

The ten `qualification-before` JSON sidecars retain the historical decoder
alias `elapsed_seconds` for GNU time `%R`, which is the minor-fault count. The
raw `.time` files remain the correct `%M %R %F` values, and the correction
receipt preserves both the frozen pre-correction driver hash and the corrected
driver hash. The qualification-before raw aggregate, schedule, plan, binaries,
and statistics were not changed. Qualification-after, memory, heaptrack, and
native `.rss` records use the correct `minor_faults` field. No elapsed-time
claim is made from the ten aliased sidecars, and qualification data is not
pooled into native estimates.

The final restoration receipt covers the 9,196 production file hashes. The
cleanup witness covers the four benchmark executable hashes, verifies them
before target removal, and records that the target is absent afterward. No
live-executable claim is made. The probe and candidate are therefore not part
of the restored production state. This review adds only this packet-local
markdown record.

## Disposition

The packet is suitable as diagnostic evidence: custody, deterministic payload
verification, report cardinality, finite CPU-task/permit/source accounting,
native pairing, snapshot identity, and heaptrack provenance all replay. It is
not an attribution or adoption result. The unresolved next measurement is a
controlled reconciliation of GNU time maximum RSS with child residency, while
separating code layout, allocator scheduling, and transient peaks. Until that
measurement exists, the rejected 0787 candidate remains rejected.
