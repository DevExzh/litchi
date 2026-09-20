# 0722 paired measurements

All deltas are candidate / baseline minus one. Negative latency deltas are faster. Times are microseconds. No pair or sample is omitted.

## Primary native

| Pair | Corpus | Phase | Baseline p50 | Candidate p50 | p50 delta | Mean delta | p95 delta | p99 delta |
| --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| pair-1 | generated | edit | 348.24 | 291.97 | -16.16% | -16.13% | -15.28% | -15.71% |
| pair-1 | generated | lifecycle | 5729.59 | 5637.55 | -1.61% | -1.56% | -1.59% | -0.75% |
| pair-1 | numbered-list | edit | 144.37 | 135.11 | -6.41% | -6.66% | -5.89% | -5.91% |
| pair-1 | numbered-list | lifecycle | 5664.44 | 5675.72 | +0.20% | -0.00% | +0.08% | -2.58% |
| pair-2 | generated | edit | 348.99 | 290.60 | -16.73% | -16.86% | -17.22% | -17.26% |
| pair-2 | generated | lifecycle | 5721.76 | 5655.22 | -1.16% | -1.18% | -1.17% | -0.26% |
| pair-2 | numbered-list | edit | 144.68 | 134.84 | -6.80% | -7.16% | -8.08% | -8.00% |
| pair-2 | numbered-list | lifecycle | 5687.12 | 5685.20 | -0.03% | -0.29% | +0.33% | -1.14% |
| pair-3 | generated | edit | 358.54 | 289.28 | -19.32% | -18.56% | -18.49% | -19.67% |
| pair-3 | generated | lifecycle | 5716.75 | 5544.49 | -3.01% | -3.04% | -2.76% | -2.58% |
| pair-3 | numbered-list | edit | 144.15 | 135.16 | -6.23% | -6.83% | -9.04% | -8.52% |
| pair-3 | numbered-list | lifecycle | 5723.48 | 5665.85 | -1.01% | -1.13% | -1.88% | -4.48% |
| pair-4 | generated | edit | 350.63 | 287.82 | -17.91% | -17.76% | -17.26% | -19.11% |
| pair-4 | generated | lifecycle | 5617.04 | 5550.69 | -1.18% | -1.39% | -0.83% | +0.96% |
| pair-4 | numbered-list | edit | 143.97 | 134.87 | -6.33% | -6.30% | -4.93% | -2.29% |
| pair-4 | numbered-list | lifecycle | 5709.75 | 5656.71 | -0.93% | -4.25% | -17.57% | -33.74% |

## Read controls

| Pair | Control | Baseline p50 | Candidate p50 | p50 delta | Mean delta | p95 delta | p99 delta |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| pair-1 | generated-medium-list-paragraphs | 65.46 | 65.42 | -0.06% | -0.20% | -0.78% | +0.60% |
| pair-1 | pinned-media-eager-paragraph-count | 101.49 | 100.65 | -0.83% | -0.45% | -1.35% | -0.25% |
| pair-2 | generated-medium-list-paragraphs | 65.46 | 65.02 | -0.67% | -0.51% | -0.21% | +0.54% |
| pair-2 | pinned-media-eager-paragraph-count | 101.57 | 100.86 | -0.69% | -0.23% | +0.53% | +1.46% |
| pair-3 | generated-medium-list-paragraphs | 65.55 | 65.32 | -0.35% | -0.18% | +0.07% | -0.01% |
| pair-3 | pinned-media-eager-paragraph-count | 102.36 | 100.61 | -1.72% | -0.58% | +0.17% | +2.60% |
| pair-4 | generated-medium-list-paragraphs | 65.87 | 65.55 | -0.49% | -0.34% | -0.04% | -3.54% |
| pair-4 | pinned-media-eager-paragraph-count | 101.24 | 100.49 | -0.74% | +0.16% | +1.55% | +18.58% |

## Allocator per-operation diagnostics

Instrumented counters are separate from native latency. Values below are medians of three samples; derived peak-above-start and net-live values are per operation and are never summed.

| Pair | Corpus | Phase | Requests baseline → candidate | Requested bytes baseline → candidate | Peak above start baseline → candidate | Net live baseline → candidate |
| --- | --- | --- | ---: | ---: | ---: | ---: |
| pair-1 | generated | edit | 8277 → 8264 | 1768329 → 1766562 | 390629 → 390629 | 388061 → 388061 |
| pair-1 | generated | lifecycle | 15616 → 15603 | 3199408 → 3197641 | 1049866 → 1049866 | 461654 → 461654 |
| pair-1 | numbered-list | edit | 2556 → 2538 | 436173 → 431414 | 21310 → 21310 | 18976 → 18976 |
| pair-1 | numbered-list | lifecycle | 4674 → 4656 | 2661485 → 2656726 | 710663 → 710663 | 122437 → 122437 |
| pair-2 | generated | edit | 8277 → 8264 | 1768329 → 1766562 | 390629 → 390629 | 388061 → 388061 |
| pair-2 | generated | lifecycle | 15616 → 15603 | 3199408 → 3197641 | 1049866 → 1049866 | 461654 → 461654 |
| pair-2 | numbered-list | edit | 2556 → 2538 | 436173 → 431414 | 21310 → 21310 | 18976 → 18976 |
| pair-2 | numbered-list | lifecycle | 4674 → 4656 | 2661485 → 2656726 | 710663 → 710663 | 122437 → 122437 |

## Failed hard gates

primary: 0 failed of 96.

read: 0 failed of 16.
