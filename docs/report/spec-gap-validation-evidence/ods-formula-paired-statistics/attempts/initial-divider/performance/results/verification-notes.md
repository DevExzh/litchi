# First capture verification and disposition

The authorized first paired-statistics capture is complete and retained under `results/`. It used baseline `aa48eee68cb6ee0904523394d4ff8015dce6e595`, the frozen candidate checkout, three warmups, fifteen fresh child processes, both `evaluate` and `parse-evaluate` phases, 39 matched controls, and 105 candidate cases. The raw receipts contain 1,170 baseline rows and 3,150 candidate rows. The candidate-wide preflight covered all 105 cases before baseline timing; its exact read receipts remain under `preflight-before-timing/`.

The profile verifier passed all raw identity, numerical, source-custody, exact read-bound, allocation-balance, and shape checks:

```text
python3 performance/verify.py --results performance/results --candidate-root /tmp/litchi-ods-paired-gates-ckzigjh1 --candidate-freeze gates/freeze.json --expected-warmups 3 --expected-samples 15
```

It reported 1,170 baseline records, 3,150 candidate records, and `status=ok`. The generated `performance-report.md` and `performance-report.json` are derived from the retained raw receipts.

The matched-control review retains ten elapsed-time triggers above the 5% review threshold. They are concentrated in the shared centered-moment consumers:

| control family | phases | baseline → candidate p50 ns/repeat | observed delta |
| --- | --- | ---: | ---: |
| sensitive inline AVEDEV | evaluate / parse-evaluate | 3,510 → 4,850; 4,040 → 5,450 | +38.18%; +34.90% |
| extreme inline AVEDEV | evaluate / parse-evaluate | 3,780 → 5,010; 4,400 → 5,650 | +32.54%; +28.41% |
| sensitive inline DEVSQ | evaluate / parse-evaluate | 3,220 → 5,460; 3,800 → 6,020 | +69.57%; +58.42% |
| extreme inline DEVSQ | evaluate / parse-evaluate | 4,110 → 6,120; 4,810 → 6,710 | +48.91%; +39.50% |
| reference DEVSQ | evaluate / parse-evaluate | 18,242 → 20,050; 18,632 → 20,375 | +9.91%; +9.35% |

`threshold-review.json` records each raw sample and receipt path. Work, resolver reads, checksums, allocation counts, requested/released bytes, peak live state, and RSS did not trigger a matched-control threshold. Candidate-only paired workloads remain absolute measurements because the baseline has no paired implementation. These ten rows are retained as investigation triggers and do not establish causal attribution or a product regression. No timing was rerun or discarded.

The source-custody snapshots are `baseline-aa48eee68cb6ee0904523394d4ff8015dce6e595/source-manifest.json` and `candidate-final/source-manifest.json`; each records unchanged selected and workspace source closures across its capture. The frozen candidate hash and gate receipts are unchanged. The root agent may archive this complete first capture and source reconstruction before any shared-helper edit.

The baseline target and candidate target cleanup receipts both record `removed: true`; the preflight target cleanup receipt also records `removed: true`. No production source or frozen profile input was edited during capture or post-capture review.
