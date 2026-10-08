# 0970 — historical baseline affinity diagnostic

**Status: descriptive diagnostic complete; no qualification or adoption.** The
96-child capture completed, the independent raw audit passes, and reducer v6
has passed independent replay and arithmetic review. Nothing here is a
candidate result, speedup claim, source attribution, or host-root-cause
finding.

## Scope and measurement boundary

All arms used the same retained historical 0967 ordinary baseline binary
(SHA-256
`0db34357a1e725dc397a4fd895532433db1214c5862b7bc96bf23f21c0396ecb`). No
candidate binary, production source edit, build, or native rerun was introduced.
The matrix has four cells — `docnohf/e1/default`, `docnohf/e32/default`,
`docfloat/e1/default`, and `docfloat/e32/default` — with 12 pairs and 24
serial children per cell. Odd pairs run unpinned then pinned; even pairs run
pinned then unpinned. Each child has five warmups and 80 measured samples,
with no retry.

The audited full mask is CPUs 0–31; pinned children use exactly CPU 31 and
unpinned children use the full allowed mask. `perf` measures the launched GNU
`time` command and inherited launcher/baseline process tree, including wrapper
startup. GNU `time` likewise includes launcher startup and the execed baseline.
The harness `Instant` samples cover only the operation-local timed region and
exclude wrapper startup and sidecar I/O. These scopes are separate and are not
interchangeable metrics.

## Capture and audit

The stage has 96/96 terminal packets, exit code 0, zero retries, 7,680 sample
rows, 576 perf-event rows, and 582 exactly referenced capture files. All
`96 × 20` report gates, `96 × 80 × 5` sample gates, and `96 × 3 × 12` route
oracle gates are true with empty failure reasons. The independent raw audit
did not import or execute the shared reducer.

The first shared reducer replay stopped on `KeyError: 'pair'` because the
packet stores pair/order fields under `context`; it did not mutate the capture.
Reducer v6 applies the reviewed context adapter, explicit within-child CV,
nearest-rank percentiles, and paired bootstrap definitions. Independent review
recomputed 96 children, 16 between summaries, and 56 bootstrap groups with a
byte-identical output and no blockers. The earlier raw-audit receipt retains
its historical `pass_with_tooling_pending` status; this v6 approval resolves
that tooling issue without changing the raw capture.

## Descriptive child-p50 result

The CV columns report sample standard deviation divided by the mean across the
12 child p50 elapsed values in each arm. The final reducer separately defines
between-child CV as population standard deviation divided by the mean across
the 12 child p50 centers, and within-child CV as the median of each child's
population standard deviation divided by its mean over 80 samples. The paired
p50 ratio is the median of pinned divided by unpinned for each pair; its interval
is a 10,000-resample paired bootstrap using seed `970001`. All values are
descriptive and do not qualify an effect.

| Cell | Pinned child-p50 CV | Unpinned child-p50 CV | Pinned/unpinned p50 ratio [95% bootstrap CI] |
| --- | ---: | ---: | ---: |
| `FloatingPictures/e1/default` | 10.3561% | 9.52625% | 0.943838 [0.864919, 1.061762] |
| `FloatingPictures/e32/default` | 1.49012% | 1.19979% | 0.999429 [0.988469, 1.001782] |
| `NoHeadFoot/e1/default` | 0.777451% | 0.823071% | 1.001952 [0.992972, 1.004463] |
| `NoHeadFoot/e32/default` | 0.476417% | 0.647617% | 0.992554 [0.991235, 0.996702] |

Pinning to CPU 31 did not remove the large `FloatingPictures/e1` dispersion;
the pinned CV is slightly higher. In that cell, Pearson association of the
12 child p50 values with whole-process instructions is 0.955 pinned and
0.892 unpinned; cycles are 0.954 and 0.909. These within-cell associations
are descriptive only. Across the 48 pinned packets, CPU migrations total
zero; 8 of 48 unpinned packets have a nonzero count, totaling 11. That is a
lead for process-level work variation, not evidence of a source, frequency,
or host root cause.

## Disposition and next diagnostic

This historical baseline packet does not qualify a candidate and launches no
candidate workload. Production is unchanged. The next step is a measured,
separately approved process-work attribution design. No further reducer work
or native rerun is implied; any attribution must still distinguish the
operation-local `Instant` region from wrapper and process-tree telemetry. The
bounded cleanup freed 1,830,912 allocated bytes and preserved the binary, raw
inventory, controls, and v6 output.

## Evidence

| Evidence | SHA-256 |
| --- | --- |
| [capture manifest](results/change-0970/capture/capture-manifest-v2.json) | `b0943c4572304dccc3a70dde17cce3894a35535e4c52fb7efb2526eb760b8685` |
| [capture stage](results/change-0970/capture/capture-stage-v2.json) | `244e033250f303328b7c4f3b1b729c4d3548886e9153b7ecfe5a996669871143` |
| [raw audit](results/change-0970/review/raw-independent-audit-0970-v1.json) | `30e2c51dbf541fdacc15d9ac23e2528b64551f4949c27659d8df409593cdf000` |
| [audit receipt](results/change-0970/review/raw-independent-audit-receipt-0970-v1.json) | `a0289335606c81deafa601c5334e71cdd542e4137ab82241c572fa1b02a363a2` |
| [independent scalars](results/change-0970/review/raw-independent-scalars-0970-v1.json) | `c9bca1f01c575678f20632e7ecdd7ded0075e7baf7094523b22914f11abfcd70` |
| [plan](results/change-0970/plan/plan-0970-v6.json) | `4043a33e413476961feb5c4712689be9bd392f341c30d7d5a8bd73406d033e7b` |
| [independent design review](results/change-0970/review/diagnostic-independent-review-v3.json) | `c612bf37d2d80c4fb7ef9f547eb8d9a86a950e04b758c380d306a0f62a770e6d` |
| [v6 reduction](results/change-0970/analysis/affinity-diagnostic-v6-final.json) | `560e4cae2cd338fd936c57d8f3aa9ecd663aa84736ce8ea5aa431a161066c4da` |
| [v6 independent approval](results/change-0970/review/reducer-v6-independent-approval-0970-v1.json) | `3b73b06783d9a697ad846fd3cf9ff2c616a56853779db69f0b0923fdf23ee1d5` |
| [v6 arithmetic review](results/change-0970/review/independent-reducer-review-0970-v1.json) | `b2582eff6624b90eeaaa42568b2fbbb095e9d4a5319b65bda61355fd5fe7d732` |
| [v6 correction receipt](results/change-0970/receipts/reducer-correction-0970-v3.json) | `f394858ba3aa43a912cc7abb4dc6e0fbae1fc5d089c9f82ce93e0fcd5e33af1c` |
| [temporary cleanup](results/change-0970/receipts/temporary-cleanup-0970-v1.json) | `f17b2eca44d8a9111dec458cd18e26a6aed6b0d85b67f01140c4cd7cccda97cb` |
