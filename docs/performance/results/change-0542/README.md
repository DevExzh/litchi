# 0542: shared XLSX traversal refusal gate

The candidate is rejected and production restored. All four valid primary rows
passed: planning p50 improved 23.5–26.4% and workflow p50 6.9–9.1%. Ordinary
planning memory and every invalid peak gate passed. Three late-raw p50 rows
exceeded the frozen 2x baseline-valid envelope. No runtime speedup is retained.

The standalone XLSX planning guard is retained as a measured enabler. The
original proposal, cap tests, immediate-fallback supplement, formatted applied
patch, source snapshots and all raw results remain for replay. The separate
`next-candidate.patch` is unmeasured and has no admission in this batch.

- [Change record](../../changes/0542-xlsx-shared-traversal-refusal-rejected.md)
- [Protocol](protocol.md), [plan](plan.json), [allocation gates](allocation-gates.json)
- [Decision](decision.json), [native comparison](comparison.json)
- [Allocation analysis](allocation-analysis.json), [refusal guards](guard-analysis.json)
- [Applied source review](source-review.md), [static review](review.md)
- [Individual adverse review](adverse-review.json)
- [Next priority](next-priority.md), [unmeasured follow-up review](next-candidate-review.md)

`python3 -B verify.py --strict` replays the retained evidence after cleanup and
sealing. Build/capture scripts are historical reproducibility records and refuse
to overwrite receipts or a sealed bundle. Prepared instruction, hardware and
eager scripts were not executed after the refusal gate failed.

`precleanup-verification.json` and `preseal-verification.json` record their
respective checkpoints. Their pending seal fields are intentional historical
state; the final `SHA256SUMS` and strict verifier establish sealed custody.
The owned target and all retained executable copies were removed after their
hashes and final checks were recorded.
