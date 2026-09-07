# Change 0457 performance acceptance review

The final formal comparison supports retaining a narrowly scoped
bounded-working-memory enabler for the specialized source-backed ODP tail
publication path. It does not support an ordinary CRUD, `Commit`/`Patch`,
latency-speedup, result-equivalence, or total-process-memory claim.

## Evidence and comparison scope

The authoritative comparison is [`comparison-final.json`](comparison-final.json)
against [`candidate-final/summary.json`](candidate-final/summary.json). It
contains 12 lanes for each role: repeats R1/R2, normal and allocator binaries,
and tiny, medium, and large source slides. Each lane retains 30 samples after
three warmups, giving 360 candidate samples, 360 control samples, and 720
formal samples in total. The final candidate is the post capacity/lease-fix
source epoch; the earlier [`comparison.json`](comparison.json) and historical
candidate summary remain immutable evidence and are not substituted for the
final endpoints. The final source, common/ODP, strict retained-plan, native
preflight, and native output gates passed. The source corpus and append request
are bound, but the candidate and control retain different result contracts and
produce different ZIP framing. The numeric endpoint comparison therefore
remains descriptive.

The candidate lifecycle times direct source opening, bounded source scan and
proof, insertion-plan preparation, and sequential publication. The source
archive `Arc`, positional provider, submitted title/body strings, options, and
discard sink are prepared before the timed region. Digest and semantic oracles,
report assembly, and operation-owned drops are after the clock. The source-tail
publication reads the content member through four bounded passes: source scan,
candidate scan, and fresh verified readers for replay measurement and emission.

The elapsed bootstrap is limited to each row's candidate and control 30-sample
vectors. It uses 10,000 independent nonparametric resamples and nearest-rank
95% endpoints for the median delta. It has no cross-row pooling, geomean,
multiple-comparisons correction, or causal interpretation. Memory metrics do not
have a corresponding bootstrap interval.

## Memory result

The allocator-mode `region_peak_above_entry` values are the callback-observed
allocator high-water minus live bytes at region entry. They are not process RSS,
physical memory, or a complete accounting of source ownership already live at
entry. The retained p50 values are:

| Source slides | Control region peak above entry | Candidate region peak above entry | Control allocated bytes | Candidate allocated bytes |
| ---: | ---: | ---: | ---: | ---: |
| 64 | 781,342 | 620,381 | 10,613,383 | 2,607,922 |
| 4,096 | 18,027,568 | 620,385 | 110,226,105 | 36,593,937 |
| 8,192 | 35,958,388 | 620,385 | 211,442,207 | 71,164,177 |

These values repeat exactly in R1 and R2. The candidate's observed payload
region is therefore flat at about 606 KiB across the three source sizes in this
matrix, while allocation volume and allocation calls still grow with document
size. This supports a scoped bounded-working-memory observation for this
specialized path when stated together with its explicit parameters. The
implementation bounds XML input, token bytes, depth, events, attributes, and
namespace/attribute parser state; ZIP replay bounds decoded bytes, compressed
bytes, and total archive bytes; and replay uses fixed 64 KiB input and output
buffers alongside preservation's fixed source-copy buffer. The resulting
working-memory envelope is bounded by those archive, token, depth, metadata,
and authored-fragment limits. It is not a claim that all memory or CPU work is
constant, nor that the whole process has a bounded or reduced RSS peak.

The source `Arc` and fixture ownership are already live at region entry. The
candidate writes to a discard sink, while the owned control retains its source
and materialized commit/output through endpoint sampling. Those ownership and
retained-result differences are material limitations on interpreting the large
negative allocator deltas as a general memory optimization.

## Positive flags and adverse observations

There are no positive elapsed-latency flags above 5%. The largest positive
elapsed result is R1 normal medium: +3.262% at p50, +3.243% at p95, and
+3.110% at p99. Its row-local bootstrap median interval is +2,296,082 to
+2,570,058.5 ns around a +2,437,189 ns median delta, still below the 5%
acceptance threshold.

The only positive process-HWM flags above 5% are:

| Lane | Metric | p50 | p95/p99 |
| --- | --- | ---: | ---: |
| R1 normal tiny | process HWM | +1,134,592 bytes / +8.301% | +1,085,440 bytes / +7.913% |
| R1 allocator tiny | process HWM | +1,380,352 bytes / +10.184% | +1,368,064 bytes / +10.085% |
| R2 normal tiny | process HWM | +905,216 bytes / +6.517% | +831,488 bytes / +5.955% |

Here “process HWM” means the in-process `/proc` `VmHWM` sampled after the
operation. The comparison labels it as `peak_rss_bytes`, but its scope is
process-lifetime high-water after sampling, not an operation peak. These three
tiny-lane increases remain adverse observations; they are not erased by the
allocator-region result. `rss_delta_bytes` has no positive p50 flag and is a
procfs after-before observation, not a peak.

The independent GNU time maximum-RSS metric has no positive flag above 5%.
Its largest positive value is R1 normal medium, +180,224 bytes / +0.213%.
GNU time measures the maximum RSS of the whole process invocation and is a
different metric from both process HWM and allocator-region high-water.

## Disposition

Retain only the specialized allocator-region memory enabler: on the pinned
64/4,096/8,192 source-slide matrix, the source-backed path shows a flat
observed operation-region payload and lower callback allocation volume under
the explicit finite XML/archive/replay limits. State the result as evidence for
this path and matrix, with the entry-live ownership and retained-result caveats.

Do not retain an ordinary `Commit`/`Patch` or CRUD speedup claim, a general
latency claim, total-process RSS reduction claim, candidate-owned-memory claim,
constant-CPU claim, scaling claim, or output/result equivalence claim. The
source-tail candidate and owned control have different publication APIs,
retained values, timing setup, and ZIP bytes.

## Profiling diagnostic

The paired large-shape sampled-cycle profiles are retained under
`profiling/r2/runs/{candidate-large,control-large}`. The candidate profile
receipt passed. The control recorder receipt preserves an original oracle
failure because that recorder invoked a stale frame-pointer expectation; the
same 100-sample report is validated by the retained amended control oracle in
[`control-oracle-amendment.json`](profiling/control-oracle-amendment.json),
whose receipt and hashes bind the unchanged report, workload, and profile
artifacts. This amendment validates the existing capture; it does not replace
or recapture it.

The candidate profile's leading sampled-cycle symbols are SHA-256 compression
(11.00%), bounded XML name validation (9.08%), `memcmp` (6.01%), and XML start
element validation (5.42%). The control profile's leading symbols are
`memcmp` (9.67%), quick-XML attribute iteration (6.00%), `memmove` (5.76%),
and namespace-prefix resolution (5.24%). This is consistent with the
candidate's added bounded scan/proof/hash work and the control's owned-parser
work, but it is diagnostic rather than a causal CPU attribution.

The profiles sample the whole recorded binary process, including setup,
warmups, retained samples, and report assembly; the independent oracle runs
after recording. They use sampled user cycles only, with no perf-stat counters,
and retain some unavailable kernel symbols. They are also bound to the earlier
pre-fix source/binary epoch, so they must not be treated as final-candidate
performance evidence. No ordinary speedup or CPU-equivalence claim follows
from them.

This review read the immutable final and historical comparisons, candidate
summaries, source capture scope, allocator derivation, process-metric
definitions, replay/XML limit implementation, and authenticated profile
receipts. No new CPU, build, test, or timing job was run for this review.
