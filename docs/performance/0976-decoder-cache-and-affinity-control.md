# 0976 — decoder cache transition and affinity control

**Status: bounded diagnostic complete; offline evidence qualified.**
0976 examines two offline/control questions raised by the failed 0975 profile:
why the immutable 0975 child-01 replay enters a long empty-callchain tail,
and whether a harmless control can move perf recording work while keeping its
target on CPU 31. External decoder shims were compiled for the replay, but no
Office workload or benchmark build, production source change, speedup claim,
or adoption decision was made. The failed 0975 cohort is not revived.

## Decoder transition

The retained 0975 corrected stream has 6,940 samples and 3,745 empty decoded
callchains. The first empty is sample index 1,389 at timestamp
303382.496301; 3,734 of the empty samples have raw IPs in the target PF_X
interval and 11 are outside it. The first target report failure is decoder log
line 70,522 with libdw error 21, `address range overlaps an existing module`;
the associated build-id error is `No such file or directory`. This is a
decoder-coverage transition in an immutable replay, not evidence of a source
hotspot or a complete-stack guarantee.

The guarded cache-control replay produces zero empty samples for both the
0975 child and the retained 0973 control. The final v2 raw-target ASLR replay
also passes independently: each replay has 6,940 samples, a 6,940,000,000
period, zero empty samples, 951 marker samples with a 951,000,000 marker
period, zero loss, 3,289 unresolved frames, and 63,282 canonical frames. Raw
script addresses differ under ASLR; after normalization, the canonical stack
sequences are identical at digest
`1f7335d982514a9f586e0a3a001343ce7332deb650a9bf6ebae41f60f62dca3c` under
target delta `0x1000000000`. Each replay records one initial cache store,
111,448 identity-valid reuses, and zero fallback; the initial-rewrite flag is
a per-call base predicate, not an event count. These are offline replay facts,
not an adopted decoder or runtime change. The checker-only v3 validation
failure is retained for review history; the final posthoc ASLR receipt passes,
and the corrected 0973/no-op regression posthoc also passes.

The corrected final-v2 0973 regression has 6,931 samples with period
6,931,000,000, zero empty samples, 947 marker samples, 3,357 unresolved
frames, and the same eight historical lost samples; its canonical output
digest `56cf8e63a673b1eb71a49539df84525b5ab2b1afdc6430d946dab946d1b80574`
matches the legacy replay. The final-v2 no-op is byte-identical to stock at
digest `03f2d314e734608d6ea71dbdf6be1a186ca03dd0cf5f3c37f8a979f4fcf5d1c3`,
with the expected 4,540 empty samples, zero markers, and eight lost samples.
This is regression evidence, not native qualification or a production change.

## Affinity control

The v6 harmless control uses the retained `instructions:u`, period 1,000,000,
`--per-thread`, `-m 16M`, and 32 KiB DWARF settings. It observes one perf
PID/TID, 3786017, moved from CPU 31 to CPU 30, while the target PID/TID
3786018 remains on CPU 31. Start, 13 periodic, and end checks all pass at a
nominal 20 ms interval. The control records 4,685 positive samples, all on
CPU 31 with period 1,000,000, raw loss zero, and terminal loss zero. The final
audit obtains the explicit `# Total Lost Samples: 0` field from
`perf report --stdio --percent-limit 100`; these remain separate loss domains.
It is a
stdlib CPU-burn control, not an Office or native workload, and its marker gate
does not apply.

The affinity v2 addendum corrects the earlier unconditional reader-thread
statement: the installed perf creates additional reader pthreads only when its
parallel-streaming option is enabled. This v6 command did not use that option
and observed one perf TID. The control therefore proves the single-TID
handoff for this exact command and host; it does not prove a future
`--threads` mode, eliminate transient changes between polls, or validate the
proposed disabled-event/control-FIFO ordering.

## Disposition

The controlled module-overlap and target-reuse replay identifies a bounded
decoder mechanism for the empty-chain transition, but does not establish why
the original path re-registers the target or identify a capture-loss, source,
or performance root cause. The affinity control narrows placement behavior
without supplying workload evidence. Preserve the raw replay and control
gates. The raw-target ASLR replay and posthoc evidence seal pass, and the
corrected 0973/no-op regression posthoc passes. Any future combined
decoder/affinity work belongs to a separately reviewed 0977 plan with its own
control before an Office workload is considered. No 0975 qualification or
production change follows.

## Evidence and review state

| Evidence | SHA-256 |
| --- | --- |
| [decoder transition analysis](results/change-0976/analysis/decoder-empty-transition-0976-v1.md) | `d0a964420c48dd5d383b02d244560ca50fb2fc2724358623fac0fbe3f73e9259` |
| [decoder transition data](results/change-0976/analysis/decoder-empty-transition-0976-v1.json) | `f3ff6bf39b47273a94f870564f4a97a40f3bd41351ca1fb56c9f425bb249355b` |
| [raw-target ASLR posthoc review](results/change-0976/receipts/aslr-cache-final-v2-0975-v3-posthoc-v2.json) | `54657f748f73e16709f795f5ceb0e76dd4d2c07afb4e487e5af2102ca2f4cac2` |
| [0973 regression/no-op posthoc review](results/change-0976/receipts/final-v2-regression-0973-v2-posthoc-v3.json) | `04dea96e90eb5ab9cfc9718154ce7e49004fb6ea5dd25f1b6c9cac0d367bd621` |
| [independent decoder review](results/change-0976/analysis/independent-decoder-review-0976-v1.json) | `b3e85c60ffdf4c61d8d2350d5088fc169c4c9a691d7bb4dd8beae304390aa442` |
| [affinity design](results/change-0976/design/design-affinity-ordering-v1.md) | `5033971234f0fb449c2fc3dcbd12512a24782535696387355864ad20870e89f7` |
| [affinity v2 addendum](results/change-0976/design/design-affinity-ordering-v2-addendum.md) | `3304f2124f3ba4590acb733133f8672ab568c92c5c0c51f639711dbd54208dbc` |
| [control plan v6](results/change-0976/plans/control-plan-0976-v6.json) | `92b40f397df1aa2c3ea8f3cb78e938d9e32d250e97990ffed4478b4da70e9991` |
| [v6 helper seal](results/change-0976/receipts/helper-seal-0976-v6.json) | `1ca27c37ff5ad54d870496249b4015a2f8cddef76c3af1c44f7c1a3ffe7f1e62` |
| [independent control audit v4](results/change-0976/analysis/independent-control-audit-0976-v4.json) | `b3c8d22275c76a369acb64c9fa09f05357225126120e3cabae3ae380c5f2eb6c` |
| [explicit terminal loss report](results/change-0976/analysis/control-perf-report-percent100-0976-v1.txt) | `956155405ee747da2f2617aa497abdfbb732852a0299e752da008c3746484e62` |
| [affinity control review v2](results/change-0976/analysis/independent-affinity-control-review-v2.json) | `cc813e02fdba61bed32eb6d8adeaaf2590affe6f35e648573964acfbbef74ebb` |
| [control command](results/change-0976/control/affinity-control-v6/command-receipt.json) | `c8c0b1babcc11b04ce401573f420975071ec3146047e75adaa9c660e3976a000` |
| [control bridge](results/change-0976/control/affinity-control-v6/bridge-receipt.json) | `7fc8af88bbc5f99f4bcebf9ed9ca89e99917255abb071264b69a55c69e1a821d` |
| [control target](results/change-0976/control/affinity-control-v6/target-receipt.json) | `0dcac8748bb1e89fa6a3f012b65fa5d028ed9b85b7a286fb7b9c89bfb573abc5` |
| [control terminal](results/change-0976/control/affinity-control-v6/terminal-receipt.json) | `14c93c60bbc165188c7726d5e4cf52ca98a0f314b9a1007427f93f31d650b747` |

The control audit is bounded to this harmless stdlib run. Static and dynamic
decoder guard reviews, the raw-target ASLR replay and posthoc evidence seal,
the corrected 0973/no-op regression posthoc, and the independent decoder
review pass. This record makes no native admission or production
recommendation.
