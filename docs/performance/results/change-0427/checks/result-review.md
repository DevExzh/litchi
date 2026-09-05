# Change 0427 retained-result review

This is an independent read-only review of the retained release evidence. I
read the four representative R1 reports, the corresponding R2 identities, all
derived tables, the report summary, capture receipts, resource logs, protocol,
and the frozen retention source. I did not run the replay or summary scripts,
builds, tests, profilers, or CPU workloads.

The evidence is internally consistent with the declared callback-order scope.
The eight receipts describe four API/corpus pairs with two fresh processes per
pair, 30 retained rows and three warmups per process. Every report identifies
source revision `392a11a1e7fca51ec930adff024b0f7a7f20d4fd`, binary
`e48043229d70195d3394631ccbc796f6155747dc070f4ea29963a6d0991a25a1`, and
`serialized_region_peak_v3`. The plain and media source/destination archive
identities match between owned and source-backed reports. Expected output
identities differ by API as permitted, and are stable across each API/corpus
repeat pair.

The representative raw rows show `measured` retention probes, zero
`overflowed` and `observer_invalid` flags, and source-backed Arc counts of
exactly one for both caller sources after document-handle drop. The retained
reports all set both correctness gates and the 33-iteration check count. The
recorded capture/replay result covers 264 exact-output comparisons, including
24 warmups, and 240 retained rows. Every final sink-drop point returns to its
own baseline live-byte value.

## Checkpoint arithmetic

The following values are live-byte changes from each row's
`baseline_before_inputs` point. They reproduce the resource review and the raw
reports; the region peak is the whole-probe peak above that same baseline.

| Corpus/API | Entry B | Published +B | Document-drop +B | Caller-drop +B | Final +B | Whole-probe peak +B |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| plain / owned | 184,507 | 618,915 | 128,626 | n/a | 0 | 1,181,495 |
| plain / source-backed | 154,519 | 277,062 | 192,119 | 128,626 | 0 | 909,874 |
| media-rich / owned | 68,932,380 | 168,526,835 | 67,265,282 | n/a | 0 | 203,809,932 |
| media-rich / source-backed | 68,888,104 | 124,056,121 | 100,897,023 | 67,265,282 | 0 | 143,640,102 |

The `Documents dropped` and `Caller-drop` columns are point values relative to
entry, as stated by the resource review. They are not the bytes released by
that individual drop. The actual caller-Arc release is
`192,119 - 128,626 = 63,493 B` for plain and
`100,897,023 - 67,265,282 = 33,631,741 B` for media-rich. Those residual values
then equal the bounded sink capacity, and the final sink drop removes that
residual. This agrees with the frozen source: source-backed publication
consumes the editor, then the plan and source view are dropped, then the two
caller Arcs, then the sink; the owned path drops snapshots and packages before
the sink.

The sink capacities also explain the residuals without assigning ownership to
global counters. Plain owned output is 31,545 B, giving the configured ceiling
`2 * 31,545 + 65,536 = 128,626 B`. Media-rich owned output is 33,599,873 B,
giving `2 * 33,599,873 + 65,536 = 67,265,282 B`. Source-backed output is
within the matching owned ceiling in both corpora. The source-backed and owned
publication points therefore retain different owner sets and must not be
presented as a memory-reduction comparison.

The whole-probe peaks are also scoped correctly. `allocation_metrics::begin()`
runs before the baseline, while input clones and bounded sink reservation are
inside the region. The peak therefore includes setup and concurrent process
callback state; it is not the historical operation-only peak and is not
object-owned memory. The `/usr/bin/time -v` resource logs cover the entire
child process, including corpus construction, gates, binary hashing and JSON
serialization. Their RSS and elapsed values remain context and do not support
a phase-local resource or performance claim.

## Repeat and claim review

The derived summary reports identical phase changes for every R1/R2 pair,
including the final caller-Arc and sink boundaries. The MiB table uses the
declared conversion from bytes; for example, `618,915 / 1,048,576 = 0.590243`
and `168,526,835 / 1,048,576 = 160.719714` after rounding. The summary and
resource review retain `performance_claim: null` and explicitly leave cache,
managed-budget, near-limit, RSS-release, leak, latency, throughput and
object-owned-memory conclusions open.

One presentation detail is worth preserving when reading the resource table:
the heading `Caller Arcs dropped delta B` names the post-boundary point delta
from entry, not the release-window delta. The prose immediately below gives
the actual release-window values above, so this is a wording ambiguity rather
than an arithmetic or evidence failure.

## Review result

The retained raw reports, derived arithmetic, ownership interpretation and
scope statements agree with the frozen retention implementation and declared
protocol. I found no evidence blocker. The bundle remains suitable for the
descriptive callback-order observation only; it does not authorize a matched
optimization or broader resource claim.

Root clarification: the resource table headers now say “After” each boundary;
the following paragraph explicitly defines every endpoint delta relative to
sample entry. The release-window byte changes remain separate in the prose.
