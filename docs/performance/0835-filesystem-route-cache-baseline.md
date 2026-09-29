# 0835 — filesystem route/cache baseline

The committed repaired harness completes **72 formal reports / 2,160 measured
samples**, plus seven separate qualification reports / sixteen samples. The
strict reader and independent numerical/custody audit pass. No planned 20%
block-spread flag is triggered. This is a descriptive current-route baseline,
with no production change or before/after optimization claim.

The complete evidence is in [change-0835](results/change-0835/README.md), with
raw reports, command receipts, source and executable identities,
[analysis.json](results/change-0835/analysis.json), and
[audit.json](results/change-0835/audit.json). The baseline source is
`77adc4f1e24bdf76f5e34ed4dfc0113e25120c59`, following the
[0834 harness repair](0834-aligned-source-harness-repair.md).

## Environment and corpus

The workload runs on an AMD EPYC 9R45 host, Linux 7.0.0-1012-aws, with Rust
1.95.0 (59807616e), the system allocator, and an ordinary release build.
Commands are pinned to CPU 12 on the recorded ext4 mount; the harness reports
the statfs family as `ext2/ext3`. Pinning does not reserve the CPU or storage
exclusively. The full host record is [host.json](results/change-0835/host.json).

The OPC corpus contains four incompressible 4 MiB logical Parts and six ZIP
members. Its ordinary archive is 16,783,632 bytes, SHA-256
`a0c1af9e2c7a19148b44fc2a8c594c7a274131d74f9f042d55b487d5337cd1e6`.
The save replaces one Part. The PPTX corpus contains 200 slides, eight text
boxes per slide, and eight 2 MiB media members; the selector accesses position
100. Its ordinary archive is 17,017,139 bytes, SHA-256
`61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757`.
These are fixed synthetic corpora, not independent Office-producer certification.

Verified-cold inputs use private, proven zero EOCD-comment padding to reach a
page boundary: OPC adds 1,776 bytes and PPTX adds 1,741 bytes. Payload semantics
are unchanged, but the source route can make an extra bounded 64 KiB metadata
tail read. Warm/cold comparisons therefore include this documented layout
difference. OPC cold outputs retain route-specific raw hashes and lengths;
the source route preserves the private comment and the eager route drops it.
The proof verifies the allowed difference without normalizing published bytes.

## Protocol

Six blocks each cover six selectors in both warm and verified-cold states.
Order alternates forward/reverse across blocks. Each command uses thirty
measured fresh child processes and three untimed warmups, with a separate
priming child before every measured operation. Corpus preparation, output
verification, and the PPTX logical-read replay remain outside operation timing.
Qualification samples are not pooled into formal distributions.

Per-block p50 uses the harness's integer midpoint; p95/p99 use nearest rank.
With thirty samples, p99 is the block maximum. The table below reports the
median of six block summaries, not pooled quantiles. The JSON retains every
block's minimum, maximum, mean, quantiles, process metrics and spread.

## Results

Latency columns are milliseconds. Peak RSS is the median of six block p50s of
child process peaks, in MiB; it is not an allocation count or a live-byte delta.
Those peaks include earlier child setup and verified-cold preparation history.
The final column is the corresponding process `read_bytes` median. A zero
value does not mean zero logical reads or zero memory traffic.

| Route | Cache state | p50 ms | p95 ms | p99 ms | Peak RSS MiB | Process read_bytes |
|---|---|---:|---:|---:|---:|---:|
| OPC open / eager | warm | 9.5440 | 9.8183 | 9.9151 | 39.95 | 0 |
| OPC open / eager | cold-verified | 52.9660 | 53.3932 | 53.4924 | 41.26 | 16,785,408 |
| OPC open / source | warm | 0.1683 | 0.1815 | 0.1854 | 8.40 | 0 |
| OPC open / source | cold-verified | 4.7664 | 4.9278 | 4.9849 | 22.80 | 94,208 |
| OPC one-Part save / eager | warm | 253.2356 | 256.6279 | 258.0726 | 44.30 | 0 |
| OPC one-Part save / eager | cold-verified | 295.5588 | 299.4901 | 301.9465 | 46.83 | 16,785,408 |
| OPC one-Part save / source | warm | 77.1097 | 78.9041 | 79.7237 | 27.12 | 0 |
| OPC one-Part save / source | cold-verified | 144.9915 | 147.6649 | 148.3203 | 28.35 | 16,785,408 |
| PPTX selected slide / buffered | warm | 7.2468 | 7.4269 | 7.4837 | 28.26 | 0 |
| PPTX selected slide / buffered | cold-verified | 50.8554 | 51.3507 | 51.6305 | 29.84 | 17,018,880 |
| PPTX selected slide / source | warm | 2.9245 | 2.9605 | 2.9753 | 12.23 | 0 |
| PPTX selected slide / source | cold-verified | 13.8719 | 14.2412 | 14.4049 | 23.05 | 282,624 |

These paired ratios divide each block's eager/buffered p50 by that block's
source p50, then summarize the six ratios. They are not ratios of the table's
aggregate medians. The interval uses 10,000 bootstrap resamples of six block
ratios, seed `835083`, and sorted indices 249 and 9,749. It describes this host
and protocol; it is not evidence of a production revision's speedup.

| Operation | Cache state | Median eager/source ratio | Bootstrap 95% interval |
|---|---|---:|---:|
| OPC open | warm | 56.815345 | 55.867505–58.513070 |
| OPC open | cold-verified | 11.097071 | 10.998129–11.217581 |
| OPC one-Part save | warm | 3.283908 | 3.279588–3.288446 |
| OPC one-Part save | cold-verified | 2.037543 | 2.036467–2.039033 |
| PPTX selected-slide lifecycle | warm | 2.478104 | 2.471252–2.485900 |
| PPTX selected-slide lifecycle | cold-verified | 3.666459 | 3.640060–3.683133 |

No latency p50/p95/p99 or peak-RSS p50 block max/min ratio exceeds the
predeclared 1.2 variability threshold. This does not establish stability below
that threshold or portability to another host. No sample is removed.

## Interpretation boundaries

OPC eager open destroys its package inside the timer, while source open retains
its package through post-timer diagnostics. That lifetime asymmetry limits
interpretation of their ratio. Borrowed OPC ingress materializes four Parts in
this corpus; source open and source overlay save record zero ordinary Part
materializations. This is not a statement about every owned OPC ingress API.

The PPTX buffered route uses `Presentation::from_bytes(fs::read(source)?)`;
the source route uses `Presentation::open(source)`. Both select one slide. The
buffered label does not establish that every media Part is decoded. PPTX logical
counters are from an untimed replay, so they do not attribute measured latency.
The aligned replay retains incidental constructor payload overlap and checks
the selected query separately; it does not claim zero unrelated bytes for the
whole open operation.

Verified cold establishes observed page-cache residency and positive process
`read_bytes`, not physical-device reads. There is no allocator-instrumented
build, cross-machine result, concurrency scaling result, or Amdahl estimate in
this batch. iWork remains excluded.

## Validation and retained corrections

All 9,389 source inputs exactly match the committed final 0834 harness. The
quality record is reused with its original scope: 569 full-suite tests passed
with one ignored before two helper amendments; the final helper passed twelve
focused tests plus formatting, compilation, Clippy, rustdoc and boundary gates.
This is not another full-suite run. The fresh release build takes 529.67 seconds
and passes, as do all seven warm/cold qualification commands.

Before formal admission, the copied plan's stale statistical settings were
reconciled with the analyzer and audit. The initial plan remains archived.
The first offline qualification reader mistakenly expected fourteen children;
six single-case reports contribute twelve and the paired report contributes
four, so the correct total is sixteen. The original reader, failed receipt and
explicit correction remain retained. No native workload was repeated.

A thirty-sample tied/unsorted statistics preflight, nine reader mutation checks,
and independent source/build/plan checks pass before capture. A separate-CPU
reader check also admits the completed first block during capture. The final
reader admits all 2,160 samples, and the independent audit reproduces the raw
statistics, command chronology, descriptors and all six bootstrap intervals.
All eighty build/workload commands succeed. Of thirteen retained offline
validation commands, only the corrected initial cardinality check fails.

Cleanup removes both exclusively owned roots: 1,784 files / 1,318,938,046 logical
bytes. Qualification, full capture, mutation checks, analysis and audit all
replay after executable deletion. The three unrelated worktree files and all
normative documents remain unchanged. The final seal binds the owned paths
and their committed Git blobs.

## Next measured question

Source-backed OPC one-Part save is the largest source-route timing here:
77.11 ms warm and 144.99 ms verified-cold. Profile its changed-member
compression, raw copying and atomic publication/durability costs before
choosing an optimization. These timings alone do not identify which phase is
the bottleneck, and the default publication contract remains intact. The
broader non-iWork performance goal remains open.
