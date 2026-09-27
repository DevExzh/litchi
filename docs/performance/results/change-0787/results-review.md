# 0787 independent raw results review

## Disposition

**Rejected under the frozen adoption policy; the baseline restoration is
correct.** The retained capture is internally consistent and the correctness,
source-work, CPU-task, and permit-release witnesses pass. The candidate has 16
eligible primed latency benefits and no p50 latency guard failure, but one
whole-child RSS case satisfies both frozen rejection clauses:

`parts / small / primed / task floor 0 / width 4`: median after/before RSS
`1.056969024`, bootstrap 95% interval `[1.032109714, 1.074751696]`.

The frozen rule rejects when the median exceeds `1.05` and the bootstrap lower
endpoint exceeds `1.0`. No rerun, threshold change, or RSS exception is
warranted. The candidate is therefore not eligible for retention even though
the latency result is substantial.

This review reads the retained receipts, reports, raw audit, frozen plan and
policy, replay output, and quality records. It did not run Cargo, native
children, profilers, or hardware-counter tools, and it did not change
production source or HEAD. I ran the packet-local pure-Python
`raw_audit.py --check`; it completed with the independent 1,080-report /
22,200-sample PASS.

## Corpus and custody

The frozen matrix has 60 cases: three Part shapes (`small`, `large`, `mixed`),
fresh and primed state, task floors `0` and `65536`, and requested widths
`1/2/4/8/32`. Six native before/after blocks use 30 samples and three
warmups; two observer blocks use two samples; qualification has one sample per
case on each leg. The retained counts are:

- 720 native reports / 21,600 samples;
- 240 observer reports / 480 samples;
- 120 qualification reports / 120 samples.

The audit reconstructs 96 deterministic member payloads, checks every ordered
output and member digest, and finds stable container/source identities across
lanes. Every report receipt exits successfully and its report and RSS hashes
match the retained files. The six-block paired bootstrap uses seed `787078`,
10,000 resamples, and endpoint indexes `249` and `9749`. I independently
replayed the nearest-rank report p50 and six-block RSS ratios; the medians and
bootstrap endpoints match the retained raw audit and paired replay.

## Parity and resource boundaries

Independent pairing found exact deterministic and resource parity for 360
native pairs, 120 observer pairs, and 60 qualification pairs. Fresh observer
operations make 64 logical source calls; primed operations make zero. For all
observer and qualification reports, requested bytes equal returned bytes,
short reads are zero, the post-operation active-read count is zero, and the
request histogram sums to the logical call count. Observed peak simultaneous
reads are bounded by the requested width on each leg.

The resource snapshots retain the same CPU-task contract on both legs: fresh
`0 -> 32 -> 32` and primed `32 -> 64 -> 64` at before-operation,
after-operation, and after-drop. Every snapshot stays within its configured
limits, and worker/I/O permits are zero after drop. Output ordering, logical
bytes, source counts, and permit/CPU-task contracts therefore do not explain
the adoption rejection.

The first full replay failure was an analyzer defect: it required exact
before/after equality for `max_active` on a fresh parallel observer case.
That value is scheduler-dependent. The valid check retains each leg's value,
requires its bound by the requested width, and keeps exact parity for calls,
requested/returned bytes, histograms, short reads, and post-operation active
reads. [replay-corrections.json](replay-corrections.json)
records that measurement inputs and policy were unchanged; the final replay
and validator agree with the raw audit's sole RSS rejection.

## Latency and RSS guard

No p50 row has both a median ratio above `1.05` and a bootstrap lower endpoint
above `1.0`; the largest p50 median ratio is `1.045032748`. Sixteen eligible
primed cases at widths 2/4/8/32 meet the benefit rule, with median p50 ratios
from `0.032511391` down to `0.009732878` (about 96.75% to 99.03% lower).
This is a bounded warm, in-memory source result and does not establish cold,
remote, or whole-Office speedups.

For the sole rejecting RSS case, the six raw whole-child pairs are:

| Block | Before peak RSS (KiB) | After peak RSS (KiB) | Ratio |
| ---: | ---: | ---: | ---: |
| 0 | 4108 | 4276 | 1.040895813 |
| 1 | 4116 | 4400 | 1.068999028 |
| 2 | 4084 | 4276 | 1.047012733 |
| 3 | 4124 | 4456 | 1.080504365 |
| 4 | 4124 | 4400 | 1.066925315 |
| 5 | 4116 | 4212 | 1.023323615 |

All six after values are higher. The before values are only about 4 MiB, and
the measured RSS is whole-child peak RSS: it includes code pages, corpus
construction, priming, allocator state, verification, and teardown. It cannot
be causally attributed to the cached operation or treated as an operation
allocation measurement. That limitation does not waive the frozen guard,
whose scope explicitly is whole-child peak RSS.

The small primed floor-0 width-32 case also has a high median RSS ratio,
`1.082901432`, but its interval `[0.971750536, 1.139141067]` crosses one and
does not independently reject under the conjunctive rule. It remains visible
in the paired output. Three p95 medians and seven p99 medians exceed a 5%
median ratio; the largest are p95 `1.223238600` (mixed, primed, floor 0,
width 1) and p99 `1.360321365` (large, primed, floor 0, width 1). These are
retained tail diagnostics, not additional adoption guards.

The final quality records pass formatting, all-feature/all-target checking,
tests, warning-denied Clippy, rustdoc, and crate-boundary checks; see
[quality-0/checks.json](quality-0/checks.json).
The restored-source receipt records all 9,196 production-file hashes matching
the baseline. The candidate's low-level source result can be retained as an
experiment record, but production retention is correctly refused until a
future memory policy and attribution study address the representative
whole-child RSS regression.
