# 0956 final comparison review

This is a read-only review of the completed v4 reduction. No reduction or workload was rerun.

Bound evidence:

- `comparison-full-v4.json`: SHA-256 `05cd72d9920a715f74c9f7c96257e6c2777bd7ec589afa38ce510290120bf8be`, 3,937,963 bytes
- `comparison-summary-v4.json`: SHA-256 `74ae2817ee86804d959b87fa88977447750c52941dbaf777de593180fa7930bb`, 276,768 bytes
- root reduction receipt `quality/reduction-v4-root.json`: SHA-256 `e204ad1df93dfc53a4154f4dbd4ab182e5aa54564118c6dac97cd2707ab66c37`, return code 0
- completed packet manifest: SHA-256 `f923ebe1bb5f8a3a58458f6458b0795ea26071ebd6ea4254c624d769b489ad28`

The summary's full-output descriptor matches the recorded full reduction SHA and size. I checked all 1,154 manifest entries: every file exists beneath the packet, has the recorded size, and matches the recorded SHA-256; declared and actual logical bytes are 23,241,616.

## Primary latency result

The publication and total point ratios are below 1.05 for every case, comparison, and p50/p95/p99 quantile. The strongest medium one-percent result is:

| Comparison | Publication p50 ratio | Publication change | Total p50 ratio | Total change |
| --- | ---: | ---: | ---: | ---: |
| width 1 vs 2 | 0.662674 | -33.73% | 0.887548 | -11.25% |
| width 1 vs 4 | 0.608413 | -39.16% | 0.867268 | -13.27% |

These are native phase ratios for this explicit fresh-process harness. They do not establish production speedup or adoption. Width 1 vs 4 medium RSS is 1.060550 (+6.055%) and correctly sets the RSS review trigger; it must not be described as an allocation claim.

Dense-sparse width 1 vs 2 publication p95 has point ratio 1.027005, but its six-block bootstrap CI is `[0.997464, 1.055880]`; the upper bound exceeds the 1.05 review threshold. This tail result must remain visible even though the point ratio is below 1.05.

## All-phase and historical flags

The primary publication/total point ratios passing does not mean the all-phase checks pass. The reducer records p50 all-phase failure in five of six groups, p95 all-phase failure in all six groups, and a p99 phase review trigger in all six groups. The flagged phase pairs are:

| Group | p50 flagged phases | p95 flagged phases | p99 flagged phases |
| --- | --- | --- | --- |
| width1/2 medium one-percent | ingress, commit-drop | ingress, commit-drop | ingress, commit-drop |
| width1/2 dense-sparse | ingress | ingress, publication | ingress, publication |
| width1/2 medium one-edit | none | ingress, commit-drop | ingress, publication |
| width1/4 medium one-percent | ingress, commit-drop | ingress, commit-drop | ingress, commit-drop |
| width1/4 dense-sparse | ingress | ingress, publication | ingress |
| width1/4 medium one-edit | publication | publication, commit-drop | ingress, publication |

These are paired process-ratio flags at the named phase quantile; they are not pooled whole-workflow latency gates. The final report should disclose them rather than summarize the run as an unqualified tail pass.

The retained 0950 report also records an ordinary medium one-edit block-2 workflow p99 increase of 5.65% (`report-draft.md`, SHA-256 `05297f9b19766f6be31796aa489c5f93162eadbc659581570abc417a99d06b53`, lines 145–147). The 0956 plan explicitly disallows an absolute comparison with 0950, so the new native data does not retire that historical flag. It should remain listed as prior evidence, with the differing protocol and scope stated.

Disposition: the reduction and packet custody are internally consistent, and the proposed conclusion is accurate only if it says “primary publication/total point ratios pass descriptively” while separately disclosing the width-4 RSS trigger, dense-sparse p95 CI upper bound, all-phase flags, and retained 0950 one-edit warning. No adoption or allocation conclusion follows.
