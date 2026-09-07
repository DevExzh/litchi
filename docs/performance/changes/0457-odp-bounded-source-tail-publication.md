# 0457 — bounded existing ODP tail publication

Existing ODP tail insertion can now validate source/candidate XML and replay
the changed ZIP member without retaining complete decoded or replacement XML.
The common path uses verified member callbacks, bounded XML token windows,
source-bound insertion proofs, deterministic two-pass ZIP replay, and typed
partial-output errors. Prepared plans retain an execution-budget lease for the
authored fragment's actual capacity until drop.

The final source-tail candidate and the owned existing-append control each
retain 360 samples across two repeats, two instrumentation modes, and
64/4,096/8,192 source slides. Their result contracts differ. At 8,192 slides,
operation-region peak allocation above entry is 620,385 versus 35,958,388 bytes,
and allocated volume is 71,164,177 versus 211,442,207 bytes. Both repeats retain
those exact allocation values. Source/fixture ownership at operation entry is
excluded; process RSS and allocation volume are not constant-space claims.

Normal-binary medium/large p50 deltas span -4.221% to +3.262%. Three tiny-fixture
process-HWM comparisons exceed +5% (+8.301%, +10.184%, +6.517%); no positive
elapsed quantile or GNU-time maximum-RSS delta exceeds that threshold. The
change is retained as a scoped working-memory capability, with no ordinary
CRUD/Commit/Patch speedup or retained-result equivalence claim.

The affected full release suites pass 1,348 tests with three ignored. Retained
validation includes thirteen insertion tests, the full common/ODP suites,
strict lint and documentation gates, and final-build checks of three
synthetic output pairs and ten native fixtures. Earlier harness/OPC and two 1,000-run
ASAN fuzz results retain their exact source epochs. Initial-source sampled
profiles guide follow-up work; they do not attribute final-source API cost.

See the [evidence bundle](../results/change-0457/README.md),
[final comparison](../results/change-0457/comparison-final.json),
[code review](../results/change-0457/final-code-review.md),
[performance review](../results/change-0457/performance-review.md), and
[next work](../results/change-0457/next-work.md). The full non-iWork goal remains
open, including integration with the existing ordinary reversible-patch
lifecycle and the wider CRUD/input/output/scaling matrix.
