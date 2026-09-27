# Independent raw results review

Review status: **the frozen raw corpus and adoption guards reproduce without a
measurement-integrity blocker.** The guard result is adoption-eligible under
the frozen policy, while the coordinator retains the final integration and
disposition decision. This review makes no production or performance claim
beyond the recorded workload and host.

## Corpus and custody

The frozen matrix contains 15 cases: tiny, medium, large, vendor, and
unicode-vendor, each measured through capture, commit, and lifecycle. The raw
reports contain:

- 180 native reports: six alternating before/after blocks, 30 samples and
  three warmups per report;
- 60 allocation reports: two before/after blocks and three samples per report;
- 15 qualification reports: one before sample for each case.

That is exactly 255 primary reports and 5,595 samples. All receipt rows exit
successfully, and the report, log, and RSS identities match their receipts.
The final raw probe reports have empty command logs. The retained failed probe
allocator attempt is separately diagnosed and is not part of the measured
matrix.

An independent raw pass found all 5,595 samples with `semantic_check` and
`reopened` true, equal expected and actual semantic text, and readback bytes
and hashes matching the published output. Source identity is stable for each
of the five shapes, and output identity is stable for each shape/mode across
legs and lanes. The two vendor shapes account for 102 reports and 2,238
samples; every vendor sample reports the six-namespace oracle over 96 text elements, each retaining six
namespaced attributes (576 attributes per presentation). The 3,357 ordinary samples correctly carry no unknown
namespace oracle.

## Latency guard

For each report I recomputed nearest-rank p50 from the raw 30-sample vector,
paired after/before by capture block, and bootstrapped the six paired ratios
with 10,000 median resamples using seed `785078`. The following values are
after/before p50 ratios; percentages are the median change and brackets are
the seeded 95% bootstrap interval.

| Case | Ratio | Change | 95% interval | Benefit gate |
| --- | ---: | ---: | ---: | --- |
| tiny/capture | 0.973265 | -2.674% | 0.968328–0.980242 | — |
| tiny/commit | 0.982316 | -1.768% | 0.975714–0.986650 | — |
| tiny/lifecycle | 0.990611 | -0.939% | 0.987493–0.995371 | — |
| medium/capture | 0.950012 | -4.999% | 0.939601–0.967519 | qualifies |
| medium/commit | 0.986009 | -1.399% | 0.982950–0.989509 | — |
| medium/lifecycle | 0.981354 | -1.865% | 0.977156–0.985238 | — |
| large/capture | 0.915107 | -8.489% | 0.910820–0.918528 | qualifies |
| large/commit | 0.981759 | -1.824% | 0.976880–0.983509 | — |
| large/lifecycle | 0.931557 | -6.844% | 0.923732–0.947058 | qualifies |
| vendor/capture | 0.957262 | -4.274% | 0.952543–0.962439 | qualifies |
| vendor/commit | 0.986703 | -1.330% | 0.982492–0.999842 | — |
| vendor/lifecycle | 0.983406 | -1.659% | 0.982994–0.986853 | — |
| unicode-vendor/capture | 0.960134 | -3.987% | 0.954889–0.968094 | qualifies |
| unicode-vendor/commit | 0.984189 | -1.581% | 0.980966–0.987111 | — |
| unicode-vendor/lifecycle | 0.983362 | -1.664% | 0.979794–0.988028 | — |

No case has a p50 ratio above `1.05` with a bootstrap lower bound above `1.0`.
Five eligible capture/lifecycle cases meet the at-least-3% improvement rule
with an upper bound below `1.0`: medium/capture, large/capture,
large/lifecycle, vendor/capture, and unicode-vendor/capture.

## Resource and spread checks

For all 90 allocation sample pairs, I recomputed
`net_live = live_after - live_before` and
`peak_above_entry = region_peak - live_before`. No after value exceeds its
before value. The audited `net_live`, `peak_above_entry`, allocation-call,
and allocated-byte fields are equal across all 90 paired samples; there are
no failed allocation calls. The two report-level p50 blocks per case also
have no resource guard violation.

The diagnostic spread audit found 22 native flags: 13 RSS flags and nine
elapsed-tail flags. The RSS flags are the six tiny groups, medium/capture
before, both legs of medium/commit and medium/lifecycle, vendor/commit before,
and unicode-vendor/capture before. The elapsed-tail flags are:

- medium/commit before p99;
- tiny/capture after p99;
- tiny/lifecycle after p95 and p99;
- vendor/capture before p95 and p99;
- vendor/commit after p99;
- unicode-vendor/commit after p99;
- unicode-vendor/lifecycle before p99.

There are no native p50 spread flags and no allocation spread flags. The two
paired p99 tail flag families are also visible directly in the raw reports:
tiny/capture block 3 rises 24.372% (260,751 to 324,302 ns), while
unicode-vendor/commit rises 6.227% in block 2 (345,912 to 367,452 ns) and
14.924% in block 5 (347,952 to 399,882 ns). These tail diagnostics do not
change the frozen p50 adoption guard.

## Assembly mechanism and limits

The assembly receipt binds the same mangled `notes::resolved` symbol in the
before and after release binaries and records hashes for both binaries and
disassemblies. The binaries were removed only under the recorded cleanup
witness; the assembly artifacts remain hash-verifiable.

The candidate assembly contains a length-indexed jump table covering 41–67
bytes. Its six destinations load the six static namespace addresses and exact
lengths—58, 46, 53, 41, 67, and 55 bytes—then converge on one `bcmp`. Equal
bytes take the static success path; a length miss or unequal comparison goes
through the existing UTF-8 fallback. This supports the intended generated-code
mechanism and also explains why unknown and same-length vendor controls remain
necessary.

The disassembly is release x86-64 evidence for code shape. It does not measure
cycles, prove causality for every public phase, or generalize to other
architectures. The paired public reports, including the vendor controls, are
the evidence used by the frozen guards.

No build, native capture, profiler, or production mutation was performed by
this reviewer.
