# 0811 current-source native results review

This review covers the terminal 0811 current-source diagnostic after all 54
native captures, two perf captures, and both decodes completed. The retained
analysis and independent raw audit report 56 reports and 1,820 probe samples;
the four deterministic gzip members retain both perf data and decoded frames.
The packet has no before/after production candidate, adoption gate, allocation
lane, or speedup claim. The current 9,196-file source map equals sealed 0810
after, so these numbers diagnose the retained source rather than compare a
new implementation.

The primary receipts are
[`analysis.json`](analysis.json) (`30355084b7669c1128d4068303bb48499332fa26f88dca89a1b2579ea05c4196`)
and [`root-audit.json`](root-audit.json)
(`09220f5c953a0813894f251fe75ad971cccb9857cb04b127154d6af171373e0d`).
The raw native timing matrix uses 30 samples after three warmups per process.
The two perf runs use the `fp` binary, CPU 12, `cycles:u` at 499 Hz, frame
pointers, 100 operations, and zero warmup. All output and semantic identities
match the sealed public-workflow oracle; production quality is reused from
0810 because the complete source map is unchanged, while the probe quality
lane freshly passes three gates and 36 tests.

## Timing perturbation and tails

The timing matrix measures two instrumentation perturbations:
`profile/control` compares the ordinary capture binary with the non-inlined
capture wrapper, and `fp/profile` adds forced frame pointers. Neither ratio is
a production before/after comparison. The paired p50 ratios and six-block
bootstrap intervals are:

| Shape | `profile/control` p50 | `fp/profile` p50 |
| --- | ---: | ---: |
| tiny | 1.0052 [1.0011, 1.0064] | 1.0255 [1.0229, 1.0367] |
| medium | 1.0096 [1.0010, 1.0145] | 1.0480 [1.0374, 1.0595] |
| large | 1.0022 [0.9995, 1.0097] | 1.0693 [1.0611, 1.0785] |

The wrapper perturbation stays near one percent or below at p50. The forced
frame-pointer perturbation grows from 2.6% on tiny to 6.9% on large, so the
perf stack lane cannot transfer its sampled native fractions to ordinary
capture latency. The large p99 `fp/profile` ratio is 1.0724 with interval
[1.0464, 1.0904]; this is still instrumentation comparison evidence.

The p99 spread flag is concentrated in the tiny fp row: 36.12% across its six
blocks, including one 339,862 ns block value against a 254,301.5 ns median.
That tail is a diagnostic outlier, not a claim about production tail latency.
The other timing spreads are small enough to describe, but no p95/p99 benefit
or regression follows from this perturbation matrix.

RSS is maximum process resident set size from `/usr/bin/time`, one value per
process rather than an allocation or live-heap series. Median RSS is 5,082,
5,292, and 18,636 KiB for ordinary tiny, medium, and large capture. The
review-only spread flags are tiny control 9.23%, tiny fp 6.53%, and medium fp
6.17%; large rows are 1.03–1.44%. Paired fp/profile RSS p50 ratios are
0.9726 [0.9430, 0.9992], 1.0105 [0.9861, 1.0594], and 0.9983 [0.9935,
1.0032] for tiny, medium, and large. These compare instrumented binaries and
do not establish an RSS or resource improvement.

## Exact-owner native stacks

Each perf repeat includes about 3,070 whole-process samples and about 960
samples qualified to the exact owner
`namespace_uri_probe::capture_region_0793`:

| Repeat | Whole-process samples | Exact-owner samples | Owner period / whole period | Lost-event lines | Unknown interior frames |
| ---: | ---: | ---: | ---: | ---: | ---: |
| 0 | 3,068 | 960 | 31.38% | 0 | 4 |
| 1 | 3,073 | 957 | 31.23% | 0 | 0 |

The owner periods are 8,470,352,078 and 8,441,181,119 sampled cycles. The
whole-process periods are 26,991,306,683 and 27,026,477,620. These are
descriptive sampled periods, not phase fractions or causal costs. The raw
audit and main analysis agree on owner counts, leaf counts, nested frame
occurrences, and paired timing values; aggregate unresolved samples are 16
and 19, with no lost-event lines.

The highest owner self-leaf periods are:

| Self leaf | Repeat 0 share | Repeat 1 share | Reading |
| --- | ---: | ---: | --- |
| `notes::scan_processed_xml` | 22.60% | 26.54% | Largest sampled leaf, but a scanner-loop owner rather than one isolated semantic operation. |
| `sha2::sha256::x86_sha::compress` | 9.28% | 9.41% | Stable native hardware-SHA leaf. |
| `Reader::read_event_impl` | 7.92% | 3.97% | Parser leaf with substantial repeat-to-repeat sampling variation. |
| `notes::inspect_element` | 5.52% | 7.42% | Element policy, attribute, resolution, UTF-8, and unescape path. |
| `core::str::converts::from_utf8` | 5.42% | 6.48% | Name/value validation inside the inspector. |
| `__memcmp_evex_movbe` | 4.79% | 5.43% | Namespace/prefix and attribute comparisons. |
| `IterState::next` | 3.33% | 4.81% | Attribute iterator work. |
| `IterState::check_for_duplicates` | 3.96% | 2.72% | Duplicate checking under attribute validation. |
| `NamespaceResolver::resolve_prefix` | 2.92% | 2.72% | Prefix lookup in the direct scanner. |

The self-leaf shares are disjoint within each exact-owner publication. Any
inclusive stack or symbol counts overlap and cannot be summed as savings.
The direct-reader wrapper remains absent from the retained scanner path, while
the required parser, namespace, duplicate, and UTF-8 checks remain present.

## Guest profile boundary

The sealed 0810 after Callgrind publications ([profile analysis](../change-0810/profile-analysis.json))
attributed 207,605,811 guest Ir, or 38.648%, to
`sha2::sha256::compress256`. That is Valgrind's software SHA path; the
established native dispatch caveat ([0807 capture profile](../../0807-pptx-current-capture-profile.md))
says Valgrind masks the SHA CPUID feature. The 0811 native owner frames instead show
`sha2::sha256::x86_sha::compress` at 9.28–9.41% of sampled owner period.
Those values are different execution mechanisms and cannot be converted into
a native saving estimate or a reason to remove a semantic digest proof.

The 0810 guest rows for `Reader::read_event_impl` (6.472%),
`inspect_element` (5.396%), `from_utf8` (4.384%), `IterState::next` (6.537%),
and `resolve_prefix` (3.559%) remain useful mechanism hints. The 0811 native
leaf distribution confirms XML scanner and attribute/name work is present, but
it does not turn guest-Ir shares into cycle shares. No historical timing pool
is used here.

## Next bounded measurement hypothesis

The next measured hypothesis should be a bounded event-local shared walk for
namespace declaration discovery and the notes attribute policy inside the
PPTX notes codec. The native leaves for `inspect_element`, `from_utf8`,
attribute iteration, duplicate checking, and prefix resolution are present in
both repeats, and the 0810 edge evidence already showed a namespace push walk
followed by checked-attribute inspection. A future experiment can test whether
one validated attribute traversal preserves the same work while reducing
those repeated visits.

This is a measurement target, not an implementation decision. It must retain
resolver push/pop timing, duplicate and reserved-name checks, unknown-prefix
refusals, UTF-8 and unescape ordering, attribute-byte accounting, XML limits,
cancellation, and exact refusal precedence. The differential oracle must
compare outputs, semantic readback, refusal identity, and preserved bytes
before any native comparison. Fresh native p50/p95/p99, RSS, owner stacks, and
allocation/resource evidence are required; a lower iterator count alone is
not evidence of benefit.

The scanner loop is the largest sampled leaf, and hardware SHA is the next
stable leaf, so the profile does not isolate a guaranteed winner between loop
overhead and XML policy work. If a line-level or owner-scoped follow-up shows
that the scanner loop dominates after the attribute walk is accounted for,
the correct result is explicit uncertainty and no candidate. A low-cost fusion
of `root_name_from_xml` and `presentation_name` should not be adopted from
this packet: those projections were not isolated by this native diagnosis,
and fusing them risks changing proof publication and error ordering without a
measured end-to-end opportunity.

The current source remains unchanged and no production candidate is proposed
by this review.
