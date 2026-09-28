# 0829 — current PPTX edit CPU phases

Fresh quality, builds, corrected address-bounded symbol admission and all
captures pass: **23 reports / 4,549 samples**. Production is unchanged.
The sealed 0828 failure remains intact; no old executable or timing is reused.

Exact edit-owner samples total 2,670/2,685. Capture has 1,196/1,193,
set-text 607/612, commit/application 867/879, and unclassified 0/1 samples.
The sample partition is independently verified and is not a wall-clock or
causal fraction. Wrapper/direct p50 ratio is 1.000715 [0.996008–1.003981];
frame-pointer/wrapped is 1.023752 [1.020221–1.027284].

See [the report](../../0829-pptx-edit-phase-profile.md), `analysis.json`,
`root-audit.json`, `frame-audit.json`, `stack-diagnostics.json` and
`results-review.md`. All reader attempts retain their source snapshots and logs.
Raw perf data and decoded frames remain in lossless deterministic gzip.

Root owns all execution and Git. Five fresh probe gates pass with three tests;
0827 production/tool quality is verified historical reuse. Target removal and
final validation are recorded in `cleanup.json` and the final reader receipt.
The three unrelated workspace files remain untouched.

Replay from the repository root:

```sh
python3 -B docs/performance/results/change-0829/validate.py --final
python3 -B docs/performance/results/change-0829/seal.py --check-head
```
