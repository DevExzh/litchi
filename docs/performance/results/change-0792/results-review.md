# Independent numeric review — change 0792

Status: **PASS; retain the production change under the frozen adoption policy.** This review covers the sealed raw reports, `root-audit.json`, `analysis.json`, the updated 0792 report, and the post-cleanup pair audit. It is a read-only numeric and custody review; it does not add a measurement or change the policy.

## Report and sample counts

| Lane | Reports | Samples | Layout |
| --- | ---: | ---: | --- |
| Qualification | 15 | 15 | One baseline sample per case; outside the paired timing result |
| Native | 180 | 5,400 | 6 blocks × 15 cases × before/after, 30 samples per report |
| Allocation | 60 | 180 | 2 blocks × 15 cases × before/after, 3 samples per report |
| **Total** | **255** | **5,595** | |

The raw reports have consistent schemas, sequential sample indices, matching source hashes, and matching output byte/hash and semantic-reopen identities. The 15-case fixture-parity check is exact. Across 120 before/after report pairs (90 native and 30 allocation), the semantic/output identity audit found zero mismatches. The separate post-capture allocation audit checks all 90 sample pairs and reports identical calls, allocated bytes, net live bytes, and peak-above-entry values in every pair.

The final quality result is 1,238 passed, 3 ignored, and 0 failed across 85 suites. All six quality gates pass. The test-only `notes/mod.rs` correction is included in the sealed source allowlist; the runtime candidate remains the empty-attribute-tail branch in `notes/codec.rs`.

## Recomputed native result

I independently recomputed nearest-rank p50 values and the paired median-ratio bootstrap from the raw native reports. The recomputation used seed `792079`, 10,000 resamples, and inclusive sorted endpoints 250 and 9749. All 15 rows matched `root-audit.json` before display rounding. Times below are milliseconds; confidence intervals are ratio intervals.

| Shape | Mode | Before p50 | After p50 | Change | 95% ratio CI |
| --- | --- | ---: | ---: | ---: | --- |
| tiny | capture | 0.237761 | 0.235631 | -0.986% | [0.988965, 0.992098] |
| tiny | commit | 0.212841 | 0.211546 | -0.672% | [0.992420, 0.998799] |
| tiny | lifecycle | 1.429182 | 1.408677 | -1.362% | [0.983237, 0.988907] |
| medium | capture | 0.469287 | 0.459647 | -1.892% | [0.969915, 0.997075] |
| medium | commit | 0.297222 | 0.295662 | -0.372% | [0.989468, 1.000922] |
| medium | lifecycle | 2.053304 | 2.012949 | -2.079% | [0.978613, 0.981792] |
| large | capture | 20.128215 | 19.091318 | -5.211% | [0.919250, 0.953105] |
| large | commit | 1.308701 | 1.294546 | -1.021% | [0.981030, 0.997316] |
| large | lifecycle | 30.768311 | 28.629127 | -6.937% | [0.925160, 0.933238] |
| vendor | capture | 0.550223 | 0.539892 | -1.810% | [0.977600, 0.986283] |
| vendor | commit | 0.326057 | 0.324597 | -0.511% | [0.993195, 0.997114] |
| vendor | lifecycle | 2.196005 | 2.163440 | -1.478% | [0.982057, 0.986709] |
| unicode | capture | 0.553533 | 0.542337 | -2.148% | [0.973081, 0.983740] |
| unicode | commit | 0.326497 | 0.323377 | -0.922% | [0.981377, 0.993791] |
| unicode | lifecycle | 2.206865 | 2.165075 | -1.878% | [0.960814, 0.982423] |

The two eligible benefits are:

- `large/capture`: ratio `0.9478948095865118`, improvement `5.210519%`, CI `[0.9192496864184769, 0.9531054964262482]`.
- `large/lifecycle`: ratio `0.9306254520711371`, improvement `6.937455%`, CI `[0.925159831067738, 0.9332382343143297]`.

Both exceed the frozen 3% useful-benefit threshold and have a bootstrap upper ratio below 1.0. No native latency guard violation is present. The 30 paired allocation block medians are unchanged for allocation calls, allocated bytes, net live bytes, and peak live bytes; the resource guard has no violation. This supports retention while making no allocation-reduction claim.

## Spread and tail limits

The analyzer reports 21 process-spread flags: 11 RSS and 10 elapsed. Six metric families have an individual paired-block change over 5%; these are spread diagnostics, not policy guard failures. Their paired median changes are:

| Group | Metric | Paired median change |
| --- | --- | ---: |
| medium/capture | RSS | +0.644% |
| medium/commit | elapsed p99 | +0.089% |
| medium/lifecycle | RSS | +0.127% |
| tiny/capture | RSS | +0.783% |
| tiny/commit | elapsed p99 | -1.391% |
| vendor/capture | elapsed p99 | -1.813% |

The vendor-capture p99 ratio interval is `[0.861128, 1.183342]`; it is retained as a spread observation and does not support a tail-improvement claim. The large baseline capture p50 process spread is 7.257%. These observations do not alter the paired p50 decision, and there is no fresh historical timing pool or guest/profile causal claim.

## Disposition and evidence boundary

The frozen policy was recorded before the build and is unchanged: at least one named capture or lifecycle case must improve by 3% with bootstrap high below 1.0, any latency guard violation rejects, and the four allocation guards may not increase. The two large cases satisfy the benefit condition, while latency and resource guards pass, so `root-audit.json`, `analysis.json`, and `disposition.json` consistently support `retained`.

The result applies to the named generated warm borrowed-input workflows and the 15 sealed cases. It does not establish a cumulative 0785 gain, a guest instruction or phase attribution, cold-start/concurrency behavior, or a broader workload result. Cleanup removed the four verified measurement binaries and the target directory; the sealed reports and replay witnesses remain available for audit.
