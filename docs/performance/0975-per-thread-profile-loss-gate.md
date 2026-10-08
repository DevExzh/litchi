# 0975 — per-thread profile pilot stops on child-02 raw loss

**Status: native diagnostic stopped; independent loss audit complete.** The
0975 batch retested the per-thread `-m 16M` path after 0974's setup-only
probes. It used the retained target and fixture with `instructions:u`, period
1,000,000, a 32 KiB DWARF stack, CPU 31, five warmups, 80 measured samples,
and a zero-loss, minimum-100-marker, stop-without-retry gate. No production
source optimization, adoption, or performance claim is made.

## Capture result

Child 01 completed with native return code 0, raw lost-record sum 0, and
terminal lost-sample value 0. Its raw parser contains 6,940 sample records;
the audited corrected decode reports 6,940 samples, 6,940,000,000 total
period, 180 marker samples, 0.0259366 marker-period share, and 3,745 empty
callchains. Stock and no-op decodes both report 6,940 samples, 4,434 empty
callchains, and zero markers. The empty-chain count makes this child useful for
checking capture behavior but insufficient for source attribution.

Child 02 also returned natively with code 0, but its raw stream contains 23
`PERF_RECORD_LOST` records summing to 184. The terminal
`PERF_RECORD_LOST_SAMPLES` value is also 184 and is the terminal representation
of that same loss, not an additional 184. It has 6,755 parsed sample records.
The audited corrected decode has 6,755 samples, 895 marker samples, and 129
empty callchains, but those values are descriptive only: raw loss invalidates
lifetime completeness, so no child-02 task-scope or marker result is treated
as validated. The runner stopped with `ProtocolError`, did not retry, and did
not start children 03–12.

The 16 MiB per-thread path therefore operated for one native child but did not
produce a zero-loss cohort. One passing child is not an all-12 result and does
not establish that buffer size, perf behavior, host state, or any source path
caused the earlier 0973 loss. No all-12 reduction, phase attribution, owner
fraction, speedup, or adoption decision follows.

## Disposition and audit

The root terminal records child 01 gate pass, child 02 raw-loss failure,
return code 2, no retry, and no all-12 reduction. The capture manifest is
failed on `raw PERF_RECORD_LOST sum is nonzero: 184`. The raw and terminal
loss counters are retained separately under the plan's non-additive loss
contract. Children 03–12 were not launched.

The independent audit binds the two child artifact sets, with a complete
packet JSON for child 01 and no child-02 packet JSON because the raw-loss gate
failed before packet admission. It confirms raw bounds, identities, child-01
complete lifetime task scope, and the separate raw/terminal loss accounting;
the observed-record parser pass for child 02 does not prove its lifetime task
history. A future batch must retain the zero-loss and single-PID gates; this
result does not authorize a retry or a relaxed gate.

## Evidence

| Evidence | SHA-256 |
| --- | --- |
| [v3 plan](results/change-0975/plans/plan-0975-pre-capture-v3.json) | `4873e502f09a3f4572ecd28d7545541e7cf7d27738c788d46fecb97293d36f97` |
| [v3 helper seal](results/change-0975/receipts/helper-seal-0975-v3.json) | `9af5da6061066645d1b7d8e40b712f39113f7c7b53e16c5b3c713793f5ac5687` |
| [root GO](results/change-0975/receipts/root-go-0975-v1.json) | `4b10643a983b525d9169e1a09697806bf76ba8eb7df7db2783349b88c3b23c54` |
| [root terminal](results/change-0975/receipts/root-terminal-0975-v1.json) | `57b2ac380aba4e23fc24638e30a63a326f6a935306cfc3c1c70b7bafc07ecfbe` |
| [capture manifest](results/change-0975/capture/capture-manifest-v2.json) | `8e56977349b453fc805b4a83f6b4c2d856652fcbb1eca4a63180ea14597ee128` |
| [child-01 raw summary](results/change-0975/capture/child-01/child-01.raw-stack.json) | `8892aa72490fba8937961542ac444122975ce78972322ed3c999ba78299b1adc` |
| [child-02 raw summary](results/change-0975/capture/child-02/child-02.raw-stack.json) | `474083e6891eaf37569e805cbc370eb85a886b9668574a7e6ac379966009ce11` |
| [child-01 terminal](results/change-0975/capture/child-01/child-01.terminal.json) | `dac7cbbba5a40ecf3595d800fe0704f66a6901e08d6c879f86035edd5ee06ff0` |
| [child-02 terminal](results/change-0975/capture/child-02/child-02.terminal.json) | `df0d0acd1879cb384d460e44b389f64a21a1417c7ffe80c9669fca0acde40796` |
| [pre-capture independent review](results/change-0975/receipts/independent-review-0975-v1.json) | `0720fae7e26c12dfbc8d06594b5fe1eb30b041d240ddf824808e955ee1bac3bf` |
| [independent failed-pilot audit](results/change-0975/receipts/independent-failed-pilot-audit-0975-v2.json) | `2695cb583aee152d98bc460163c6d11256ed796ba9bb2691440c63903396a75d` |
| [failed-capture analyzer](results/change-0975/receipts/failed-capture-analysis-0975-v3.json) | `9f272a9de0430db03cfee3c27f41af861dd0c7b9663ddd2dff429b1fa2f0cea3` |

**Disposition:** child-02 raw-loss failure; no retry, no all-12 attribution,
and no production change. The audit status is `failed_pilot_no_adoption`.

The read-only [next-capture feasibility note](results/change-0975/analysis/next-capture-feasibility-v3.md)
(SHA-256 `a49309d77ef3e051b2166fffba42d3a679ffe3b6af0f8047b37f117de2b4d586`)
records a bounded forward design and startup-ordering risk. It reports no
probe, implementation, or cause finding.
