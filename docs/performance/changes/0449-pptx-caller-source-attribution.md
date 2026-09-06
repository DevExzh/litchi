# 0449: Separate PPTX caller CPU and source-owner reads

Retain a portable reanalysis of the two 0448 profiles and eight reports/240
samples. The analysis changes which optimization is justified; no production
code or benchmark timing changes in this batch.

| Finding | Separate sleeps | Minimum service |
| --- | ---: | ---: |
| Untimed output hash, % lifecycle-frame SHA period | 49.762% | 50.088% |
| Planning touched digest, % lifecycle-frame period | 15.879% | 15.510% |
| Publication touched digest, % lifecycle-frame period | 16.115% | 15.930% |

The output hash runs after all API timers and owner drops. Lifecycle-frame CPU
samples also include warmups and untimed work and exclude blocked sleep. Thus
aggregate SHA dominance is not a timer-only production hotspot or SIMD mandate.
All caller stacks and unclassified SHA samples remain available.

Every media-rich publication sample returns 16,786,581 bytes from the source and
16,830,603 from the destination, with 425/833 calls respectively. Its source cache
records 23 hits, zero cold loads and unchanged 16,807,458 retained bytes. Source
logical revalidation hits cache; OPC compressed transfer authorization still
captures source bytes, while destination preservation reads untouched media.
The combined 33,617,184-byte total cannot be treated as repeated source reads.

Investigate OPC/ZIP authorization alongside first cold decode, preserving bounded
capture ownership, decoded equality, CRC, source lineage/version and cancellation.
Separately measure the existing 32 KiB destination copy granularity. Per-owner
phase counters and static callsites do not prove exactly removable reads; the
current journal has no per-member offsets or callsite tracing. See
[source review](../results/change-0449/source-review.md) and
[complete measurements](../results/change-0449/measurements.md).

Seven parser/classification checks and six exported corruption probes pass.
The strict inherited report oracle and period/count conservation replay all
inputs. A check-count receipt correction is retained in draft history. No new
Rust, workload or native test was executed. Replay without Git, the original
workspace, profiler or binary:
`python3 -B docs/performance/results/change-0449/verify.py --sealed`.

The full non-iWork goal remains active. Native breadth, cold I/O, bounded existing
append, repackaging and scaling evidence remain incomplete.
