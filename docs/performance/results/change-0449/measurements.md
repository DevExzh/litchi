# PPTX caller and source-owner attribution

Reanalysis of 0448: two retained media-rich profiles and eight reports/240
samples. No new workload ran. Percentages below use all lifecycle-frame
sample period, including warmups and untimed work; blocked sleep is absent.

| SHA caller | Separate sleeps: % run period | Minimum service: % run period |
| --- | ---: | ---: |
| untimed-harness-output-hash | 31.812 | 31.672 |
| planning-touched-digest | 15.879 | 15.510 |
| publication-touched-digest | 16.115 | 15.930 |
| unclassified-lifecycle-sha | 0.122 | 0.120 |

The untimed hash contributes 49.762%/50.088% of lifecycle SHA period.
The remaining unclassified SHA stacks are retained and resolve to publication
graph_digest. Both profiles have zero missing callchains. Full stacks and
outside-lifecycle SHA remain in attribution.json with conserved counts/period.

| Corpus | Phase | Source calls | Source bytes | Destination calls | Destination bytes |
| --- | --- | ---: | ---: | ---: | ---: |
| plain | opened | 83 | 8,993 | 79 | 8,579 |
| plain | planned | 16 | 5,112 | 20 | 5,644 |
| plain | published | 0 | 0 | 281 | 44,209 |
| media-rich | opened | 83 | 9,818 | 79 | 9,408 |
| media-rich | planned | 560 | 16,788,178 | 20 | 5,836 |
| media-rich | published | 425 | 16,786,581 | 833 | 16,830,603 |

Every source/destination read delta, histogram, nominal pacing counter and cache
point matches across all 30 samples, both repeats and both policies within its
corpus. Media-rich publication has 23 source-cache hits, zero cold loads and
16,807,458 retained cache bytes. The source remains cached while compressed
transfer authorization issues fresh physical reads. Destination preservation
also reads its original media. Total phase bytes are not all source bytes.

Static source corroborates these paths, but the reports contain no member/offset
read trace. No exact per-member attribution, removable-byte count, production
speedup, allocation reduction, actual link rate or Amdahl bound is established.
See source-review.md for ownership constraints and ranked follow-up.
