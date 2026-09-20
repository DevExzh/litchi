# 0721 paired measurements

All deltas are candidate / baseline minus one. Negative latency deltas are faster. Times are microseconds. No pair or sample is omitted.

## Primary native

| Pair | Corpus | Phase | Baseline p50 | Candidate p50 | p50 delta | Mean delta | p95 delta | p99 delta |
| --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| pair-1 | generated | edit | 352.85 | 313.02 | -11.29% | -0.68% | +20.87% | +20.36% |
| pair-1 | generated | lifecycle | 5684.21 | 5771.50 | +1.54% | +1.47% | +1.37% | +0.90% |
| pair-1 | numbered-list | edit | 143.09 | 135.04 | -5.62% | -5.49% | -4.80% | -5.05% |
| pair-1 | numbered-list | lifecycle | 5749.29 | 5764.69 | +0.27% | -1.12% | -1.89% | -11.63% |
| pair-2 | generated | edit | 351.12 | 311.59 | -11.26% | -11.89% | -11.83% | -13.50% |
| pair-2 | generated | lifecycle | 5811.30 | 5765.76 | -0.78% | -0.66% | -0.97% | +0.22% |
| pair-2 | numbered-list | edit | 142.92 | 135.76 | -5.01% | -5.11% | -5.38% | -6.08% |
| pair-2 | numbered-list | lifecycle | 5788.84 | 5767.34 | -0.37% | -0.38% | -0.14% | +0.44% |
| pair-3 | generated | edit | 353.12 | 305.79 | -13.40% | -13.35% | -13.13% | -11.82% |
| pair-3 | generated | lifecycle | 5717.13 | 5719.52 | +0.04% | +0.26% | -0.06% | +0.05% |
| pair-3 | numbered-list | edit | 144.45 | 135.43 | -6.25% | -6.00% | -5.55% | -4.56% |
| pair-3 | numbered-list | lifecycle | 5729.77 | 5725.72 | -0.07% | -0.20% | -0.45% | -0.47% |
| pair-4 | generated | edit | 356.25 | 304.54 | -14.51% | -14.78% | -16.61% | -16.78% |
| pair-4 | generated | lifecycle | 5796.85 | 5767.57 | -0.51% | -0.28% | -0.22% | -2.89% |
| pair-4 | numbered-list | edit | 143.03 | 135.49 | -5.27% | -5.50% | -5.85% | -6.29% |
| pair-4 | numbered-list | lifecycle | 5771.04 | 5787.64 | +0.29% | +1.31% | +2.37% | +15.51% |

## Read controls

| Pair | Control | Baseline p50 | Candidate p50 | p50 delta | Mean delta | p95 delta | p99 delta |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| pair-1 | generated-medium-list-paragraphs | 66.55 | 71.02 | +6.72% | +7.14% | +7.56% | +8.83% |
| pair-1 | pinned-media-eager-paragraph-count | 99.83 | 105.86 | +6.05% | +5.87% | +4.21% | +5.94% |
| pair-2 | generated-medium-list-paragraphs | 66.00 | 71.67 | +8.60% | +8.33% | +8.00% | +6.26% |
| pair-2 | pinned-media-eager-paragraph-count | 100.06 | 105.84 | +5.78% | +5.69% | +4.97% | +11.89% |
| pair-3 | generated-medium-list-paragraphs | 66.02 | 71.14 | +7.76% | +7.64% | +7.47% | +5.18% |
| pair-3 | pinned-media-eager-paragraph-count | 100.03 | 105.62 | +5.59% | +5.88% | +5.24% | +4.67% |
| pair-4 | generated-medium-list-paragraphs | 66.09 | 71.76 | +8.57% | +8.13% | +7.24% | +0.70% |
| pair-4 | pinned-media-eager-paragraph-count | 99.99 | 105.81 | +5.83% | +5.93% | +7.13% | +7.83% |

## Allocator requests

Instrumented request counts and requested bytes are separate from native latency. Values below are medians of three samples.

| Pair | Corpus | Phase | Requests baseline → candidate | Requested bytes baseline → candidate |
| --- | --- | --- | ---: | ---: |
| pair-1 | generated | edit | 8277 → 8264 | 1768329 → 1766562 |
| pair-1 | generated | lifecycle | 15616 → 15603 | 3199408 → 3197641 |
| pair-1 | numbered-list | edit | 2556 → 2538 | 436173 → 431414 |
| pair-1 | numbered-list | lifecycle | 4674 → 4656 | 2661485 → 2656726 |
| pair-2 | generated | edit | 8277 → 8264 | 1768329 → 1766562 |
| pair-2 | generated | lifecycle | 15616 → 15603 | 3199408 → 3197641 |
| pair-2 | numbered-list | edit | 2556 → 2538 | 436173 → 431414 |
| pair-2 | numbered-list | lifecycle | 4674 → 4656 | 2661485 → 2656726 |

## Failed hard gates

primary: 1 failed of 64.
- `{"name": "pair-1/generated/edit/mean/improvement", "observed_improvement_percent": 0.6840076270837492, "pass": false, "threshold_percent": 3}`

read: 16 failed of 16.
- `{"control_id": "generated-medium-list-paragraphs", "metric": "p50", "observed_delta_percent": 6.716754320060114, "pair": "pair-1", "pass": false, "threshold_percent": 3}`
- `{"control_id": "generated-medium-list-paragraphs", "metric": "mean", "observed_delta_percent": 7.137400577734376, "pair": "pair-1", "pass": false, "threshold_percent": 3}`
- `{"control_id": "pinned-media-eager-paragraph-count", "metric": "p50", "observed_delta_percent": 6.05058852992737, "pair": "pair-1", "pass": false, "threshold_percent": 3}`
- `{"control_id": "pinned-media-eager-paragraph-count", "metric": "mean", "observed_delta_percent": 5.872524927259071, "pair": "pair-1", "pass": false, "threshold_percent": 3}`
- `{"control_id": "generated-medium-list-paragraphs", "metric": "p50", "observed_delta_percent": 8.598484848484844, "pair": "pair-2", "pass": false, "threshold_percent": 3}`
- `{"control_id": "generated-medium-list-paragraphs", "metric": "mean", "observed_delta_percent": 8.32659496821253, "pair": "pair-2", "pass": false, "threshold_percent": 3}`
- `{"control_id": "pinned-media-eager-paragraph-count", "metric": "p50", "observed_delta_percent": 5.780762772847203, "pair": "pair-2", "pass": false, "threshold_percent": 3}`
- `{"control_id": "pinned-media-eager-paragraph-count", "metric": "mean", "observed_delta_percent": 5.685167171655725, "pair": "pair-2", "pass": false, "threshold_percent": 3}`
- `{"control_id": "generated-medium-list-paragraphs", "metric": "p50", "observed_delta_percent": 7.756740381702509, "pair": "pair-3", "pass": false, "threshold_percent": 3}`
- `{"control_id": "generated-medium-list-paragraphs", "metric": "mean", "observed_delta_percent": 7.638767103213251, "pair": "pair-3", "pass": false, "threshold_percent": 3}`
- `{"control_id": "pinned-media-eager-paragraph-count", "metric": "p50", "observed_delta_percent": 5.589602599350152, "pair": "pair-3", "pass": false, "threshold_percent": 3}`
- `{"control_id": "pinned-media-eager-paragraph-count", "metric": "mean", "observed_delta_percent": 5.878824095898061, "pair": "pair-3", "pass": false, "threshold_percent": 3}`
- `{"control_id": "generated-medium-list-paragraphs", "metric": "p50", "observed_delta_percent": 8.570996293214318, "pair": "pair-4", "pass": false, "threshold_percent": 3}`
- `{"control_id": "generated-medium-list-paragraphs", "metric": "mean", "observed_delta_percent": 8.13210260230035, "pair": "pair-4", "pass": false, "threshold_percent": 3}`
- `{"control_id": "pinned-media-eager-paragraph-count", "metric": "p50", "observed_delta_percent": 5.825582558255826, "pair": "pair-4", "pass": false, "threshold_percent": 3}`
- `{"control_id": "pinned-media-eager-paragraph-count", "metric": "mean", "observed_delta_percent": 5.933792276458161, "pair": "pair-4", "pass": false, "threshold_percent": 3}`
