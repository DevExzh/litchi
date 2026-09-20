# Second capture verification and disposition

The complete second paired-statistics capture is retained under `results/` after the shared dyadic helper change. It used baseline `aa48eee68cb6ee0904523394d4ff8015dce6e595`, the restaged frozen candidate checkout, three warmups, fifteen fresh child processes, both `evaluate` and `parse-evaluate` phases, 39 matched controls, and 105 candidate cases. The raw receipts contain 1,170 baseline rows and 3,150 candidate rows. Candidate-wide preflight covered all 105 cases before baseline timing, with exact read bounds retained under `preflight-before-timing/`.

The candidate freeze records the dyadic source hash `9509ed7613b305ee6ae8f7e7aa6136ad88a5450081f0f231ff0079fafce9a411`. The candidate source manifest records unchanged selected/workspace source closures across this capture, and the profile input snapshots before and after are identical.

The first frozen capture is archived separately under the root agent's initial-divider attempt. This result tree contains only the second capture; its raw source manifests and receipts are independent of that archive.

The second capture leaves three matched-control elapsed-time review triggers above 5%:

| control | phase | baseline → candidate p50 ns/repeat | observed delta |
| --- | --- | ---: | ---: |
| sensitive inline DEVSQ | evaluate | 3,210 → 3,480 | +8.41% |
| sensitive inline DEVSQ | parse-evaluate | 3,740 → 4,050 | +8.29% |
| sensitive inline SKEWP | parse-evaluate | 4,830 → 5,080 | +5.18% |

The AVEDEV and extreme DEVSQ triggers from the first capture no longer exceed the review threshold after the helper change. `threshold-review.json` records all three residual flags with all 15 raw samples and receipt paths. Work, resolver reads, checksums, allocation counts, requested/released bytes, peak live state, and RSS did not trigger a matched-control threshold. Candidate-only paired workloads remain absolute measurements because baseline has no paired implementation. The residual flags are retained for review and do not establish causal attribution; no timing was rerun or discarded.

The generated `performance-report.md` and `performance-report.json` are derived from the retained raw receipts. The capture receipt is `capture-summary.json`. Both capture targets and the candidate preflight target have `removed: true` cleanup receipts. No production source or frozen profile input was edited after this capture started.


An independent seeded bootstrap over the retained raw samples (seed `20260919`, 10,000 resamples, 15 observations per side, median ratio) gives 95% candidate-minus-baseline intervals of +5.8104%..+11.3208% for sensitive-inline DEVSQ evaluate, +5.8981%..+9.3834% for sensitive-inline DEVSQ parse-evaluate, and +2.6680%..+6.4990% for sensitive-inline SKEWP parse-evaluate. The DEVSQ direction is therefore retained as evidence of a small exact-ratio implementation tradeoff; SKEWP overlaps the threshold. The residual absolute overhead is 250–310 ns/repeat on these tiny inline fixtures. It is accepted as bounded after the material helper fix, with no allocation/work/read/RSS or large-reference latency flags, and remains a possible future optimization target. This does not establish causality.
