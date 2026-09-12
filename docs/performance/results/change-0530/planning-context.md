# 0530 XLSX planning phase context

Status: validated context report from the retained 0529 baseline.

Retained analyzer SHA-256: `1228b45f44c273666c65a9ffa4f146ccf6a11c58539f76ae60ca02fbcf3ac10e`.
Retained plan SHA-256: `17988e94a8d0993a6b453ca8c3ec3c46689bc89f2f823cace045a47f92b06a33`.
Retained baseline source manifest SHA-256: `3d511612ae1d8fa20b8da57db42f70ea70dc5712b844f7b73e62ceab6397e709`.

The current 0529 analyzer validated source identity, corpus identity, sink/output identity, phase vectors, and report statistics before this report summed the four 200-sample primary rows. Shares use sums and arithmetic means only. `reopen_ns` is reported separately because it is post-publication verification and is outside the lifecycle elapsed phase-sum denominator.

## Primary rows

| repeat | shape | samples | open mean/share | plan mean/share | commit mean/share | publication mean/share | reopen mean |
| ---: | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| 1 | dense-sparse | 200 | 92020.100 ns / 0.215746% | 13524591.020 ns / 31.709150% | 15371323.630 ns / 36.038917% | 13664078.835 ns / 32.036187% | 39923581.705 ns |
| 1 | medium | 200 | 85692.125 ns / 0.390685% | 7140436.670 ns / 32.554437% | 7977465.140 ns / 36.370589% | 6730241.640 ns / 30.684290% | 31385239.330 ns |
| 2 | dense-sparse | 200 | 108695.020 ns / 0.251790% | 13541900.705 ns / 31.369543% | 15196843.445 ns / 35.203185% | 14321505.090 ns / 33.175481% | 46149982.975 ns |
| 2 | medium | 200 | 88290.565 ns / 0.396500% | 7160493.585 ns / 32.156727% | 7946071.465 ns / 35.684642% | 7072626.770 ns / 31.762130% | 31384738.645 ns |

## Aggregate context

Across all four primary rows (800 samples), the lifecycle phase shares are:

| phase | sum | arithmetic mean | share of lifecycle elapsed |
| --- | ---: | ---: | ---: |
| `open_ns` | 74939562 ns | 93674.452 ns | 0.288180% |
| `plan_ns` | 8273484396 ns | 10341855.495 ns | 31.815642% |
| `commit_ns` | 9298340736 ns | 11622925.920 ns | 35.756722% |
| `publication_ns` | 8357690467 ns | 10447113.084 ns | 32.139456% |
| `elapsed_ns` denominator | 26004455161 ns | 32505568.951 ns | 100.000000% |

`reopen_ns` sums to 29768708531 ns (mean 37210885.664 ns) and is excluded from the lifecycle share denominator.

The aggregate share ranking is `commit_ns` 35.756722%, `publication_ns` 32.139456%, `plan_ns` 31.815642%, `open_ns` 0.288180%. These shares describe where measured lifecycle time was spent; they do not convert latency to instruction counts or make a new performance claim.

The 0529 pilot was rejected on its native timing gates, so this context does not admit a conditional profile or an OLE2/OOXML speedup claim. ODF remains deferred under the active OLE2/OOXML priority.
