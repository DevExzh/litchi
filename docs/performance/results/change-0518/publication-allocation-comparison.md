# DOCX publication allocation comparison

Status: **complete**.

descriptive allocator evidence from an isolated standalone probe; the baseline/candidate comparison is valid within that identical probe, but absolute allocator counts and peaks do not represent normal native production-binary values or the workspace native release profile; no native timing or speedup claim

Allocation scope: publish_document_commit_to_stream method only; region begins immediately before the call and finishes immediately after return, before the caller drops the returned Snapshot.
Timing scope: allocation counters are instrumented publication-region observations; JSON emission and allocator instrumentation are outside the native CSV timing claim.

The probe is bound to source `89ffdd564f6a76814fcc8b769f324557bd15a5201cb366ce9e75ade0a1c61181` (907 lines) and generated source `40f06c12173b0d5552f200cb1ef98b3e8172402094be8afae352f8b4ef9da768` (990 lines).  The canonical observer revision is `serialized_region_peak_v3` with revision hash `379667d71f4325361774239c7088619ab32e2ab468d3daa817b741892d84722c` and source hash `57e7c03432434bcbe7e0415eab08c599c1faa72638db5a663e92cfeb510c5897`.

Build scope: the allocator binary is built from the standalone probe manifest `allocator-probe/Cargo.toml` and lockfile `allocator-probe/Cargo.lock`. That lockfile resolves dependencies independently (`ryu` 1.0.23 versus workspace 1.0.22). The standalone manifest has no `[profile.release]`, so its release defaults use LTO off and panic unwinding; the workspace native release profile uses LTO=true and panic=abort.

Both allocator variants use the identical frozen probe source, standalone Cargo.toml, standalone Cargo.lock, and standalone release defaults; their baseline/candidate comparison is valid within this probe.

Do not interpret absolute allocator counts or peaks as normal native production-binary counts or as measurements under the workspace native release profile.

Each comparison uses two fresh one-row processes per case.  The absolute peak is `peak_live_bytes_after`; incremental peak is `region_peak_live_bytes - live_bytes_before`.  Full raw CSV rows and tagged JSON samples remain in the JSON report under `captures`.

| Case | Allocation calls | Reallocation calls | Allocated bytes | Deallocated bytes | Region peak | Absolute peak | Incremental peak |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `p128-k1-owned-repeated` | 3252.0 → 109.0 (-96.65%) | 36.0 → 12.0 (-66.67%) | 300922.0 → 75001.0 (-75.08%) | 305115.0 → 79194.0 (-74.04%) | 3817617.0 → 3815610.0 (-0.05%) | 3852590.0 → 3852595.0 (+0.00%) | 74547.0 → 72535.0 (-2.70%) |
| `p128-k1-owned-batch` | 3252.0 → 109.0 (-96.65%) | 36.0 → 12.0 (-66.67%) | 300922.0 → 75001.0 (-75.08%) | 305115.0 → 79194.0 (-74.04%) | 3817602.0 → 3815595.0 (-0.05%) | 3852575.0 → 3852580.0 (+0.00%) | 74547.0 → 72535.0 (-2.70%) |
| `p128-k1-file-repeated` | 3252.0 → 109.0 (-96.65%) | 36.0 → 12.0 (-66.67%) | 300922.0 → 75001.0 (-75.08%) | 305115.0 → 79194.0 (-74.04%) | 3284769.0 → 3282762.0 (-0.06%) | 3834754.0 → 3832747.0 (-0.05%) | 74547.0 → 72535.0 (-2.70%) |
| `p128-k1-file-batch` | 3252.0 → 109.0 (-96.65%) | 36.0 → 12.0 (-66.67%) | 300922.0 → 75001.0 (-75.08%) | 305115.0 → 79194.0 (-74.04%) | 3284754.0 → 3282747.0 (-0.06%) | 3834739.0 → 3832732.0 (-0.05%) | 74547.0 → 72535.0 (-2.70%) |
| `p128-k8-owned-repeated` | 3252.0 → 109.0 (-96.65%) | 36.0 → 12.0 (-66.67%) | 300922.0 → 75001.0 (-75.08%) | 305115.0 → 79194.0 (-74.04%) | 3820298.0 → 3818291.0 (-0.05%) | 3853248.0 → 3853253.0 (+0.00%) | 74547.0 → 72535.0 (-2.70%) |
| `p128-k8-owned-batch` | 3252.0 → 109.0 (-96.65%) | 36.0 → 12.0 (-66.67%) | 300922.0 → 75001.0 (-75.08%) | 305115.0 → 79194.0 (-74.04%) | 3820283.0 → 3818276.0 (-0.05%) | 3853233.0 → 3853238.0 (+0.00%) | 74547.0 → 72535.0 (-2.70%) |
| `p128-k8-file-repeated` | 3252.0 → 109.0 (-96.65%) | 36.0 → 12.0 (-66.67%) | 300922.0 → 75001.0 (-75.08%) | 305115.0 → 79194.0 (-74.04%) | 3287450.0 → 3285443.0 (-0.06%) | 3839458.0 → 3837451.0 (-0.05%) | 74547.0 → 72535.0 (-2.70%) |
| `p128-k8-file-batch` | 3252.0 → 109.0 (-96.65%) | 36.0 → 12.0 (-66.67%) | 300922.0 → 75001.0 (-75.08%) | 305115.0 → 79194.0 (-74.04%) | 3287435.0 → 3285428.0 (-0.06%) | 3839443.0 → 3837436.0 (-0.05%) | 74547.0 → 72535.0 (-2.70%) |
| `p128-k32-owned-repeated` | 3252.0 → 109.0 (-96.65%) | 36.0 → 12.0 (-66.67%) | 300922.0 → 75001.0 (-75.08%) | 305115.0 → 79194.0 (-74.04%) | 3829495.0 → 3827488.0 (-0.05%) | 3855596.0 → 3855514.0 (-0.00%) | 74547.0 → 72535.0 (-2.70%) |
| `p128-k32-owned-batch` | 3252.0 → 109.0 (-96.65%) | 36.0 → 12.0 (-66.67%) | 300922.0 → 75001.0 (-75.08%) | 305115.0 → 79194.0 (-74.04%) | 3829480.0 → 3827473.0 (-0.05%) | 3855581.0 → 3855499.0 (-0.00%) | 74547.0 → 72535.0 (-2.70%) |
| `p128-k32-file-repeated` | 3252.0 → 109.0 (-96.65%) | 36.0 → 12.0 (-66.67%) | 300922.0 → 75001.0 (-75.08%) | 305115.0 → 79194.0 (-74.04%) | 3296647.0 → 3294640.0 (-0.06%) | 3855591.0 → 3853584.0 (-0.05%) | 74547.0 → 72535.0 (-2.70%) |
| `p128-k32-file-batch` | 3252.0 → 109.0 (-96.65%) | 36.0 → 12.0 (-66.67%) | 300922.0 → 75001.0 (-75.08%) | 305115.0 → 79194.0 (-74.04%) | 3296632.0 → 3294625.0 (-0.06%) | 3855576.0 → 3853569.0 (-0.05%) | 74547.0 → 72535.0 (-2.70%) |
| `p512-k1-owned-repeated` | 12470.0 → 109.0 (-99.13%) | 38.0 → 12.0 (-68.42%) | 947578.0 → 75001.0 (-92.08%) | 951771.0 → 79194.0 (-91.68%) | 3954705.0 → 3949626.0 (-0.13%) | 4020999.0 → 4015920.0 (-0.13%) | 77619.0 → 72535.0 (-6.55%) |
| `p512-k1-owned-batch` | 12470.0 → 109.0 (-99.13%) | 38.0 → 12.0 (-68.42%) | 947578.0 → 75001.0 (-92.08%) | 951771.0 → 79194.0 (-91.68%) | 3954690.0 → 3949611.0 (-0.13%) | 4020984.0 → 4015905.0 (-0.13%) | 77619.0 → 72535.0 (-6.55%) |
| `p512-k1-file-repeated` | 12470.0 → 109.0 (-99.13%) | 38.0 → 12.0 (-68.42%) | 947578.0 → 75001.0 (-92.08%) | 951771.0 → 79194.0 (-91.68%) | 3400353.0 → 3395274.0 (-0.15%) | 4020994.0 → 4015915.0 (-0.13%) | 77619.0 → 72535.0 (-6.55%) |
| `p512-k1-file-batch` | 12470.0 → 109.0 (-99.13%) | 38.0 → 12.0 (-68.42%) | 947578.0 → 75001.0 (-92.08%) | 951771.0 → 79194.0 (-91.68%) | 3400338.0 → 3395259.0 (-0.15%) | 4020979.0 → 4015900.0 (-0.13%) | 77619.0 → 72535.0 (-6.55%) |
| `p512-k8-owned-repeated` | 12470.0 → 109.0 (-99.13%) | 38.0 → 12.0 (-68.42%) | 947578.0 → 75001.0 (-92.08%) | 951771.0 → 79194.0 (-91.68%) | 3957386.0 → 3952307.0 (-0.13%) | 4025703.0 → 4020624.0 (-0.13%) | 77619.0 → 72535.0 (-6.55%) |
| `p512-k8-owned-batch` | 12470.0 → 109.0 (-99.13%) | 38.0 → 12.0 (-68.42%) | 947578.0 → 75001.0 (-92.08%) | 951771.0 → 79194.0 (-91.68%) | 3957371.0 → 3952292.0 (-0.13%) | 4025688.0 → 4020609.0 (-0.13%) | 77619.0 → 72535.0 (-6.55%) |
| `p512-k8-file-repeated` | 12470.0 → 109.0 (-99.13%) | 38.0 → 12.0 (-68.42%) | 947578.0 → 75001.0 (-92.08%) | 951771.0 → 79194.0 (-91.68%) | 3403034.0 → 3397955.0 (-0.15%) | 4025698.0 → 4020619.0 (-0.13%) | 77619.0 → 72535.0 (-6.55%) |
| `p512-k8-file-batch` | 12470.0 → 109.0 (-99.13%) | 38.0 → 12.0 (-68.42%) | 947578.0 → 75001.0 (-92.08%) | 951771.0 → 79194.0 (-91.68%) | 3403019.0 → 3397940.0 (-0.15%) | 4025683.0 → 4020604.0 (-0.13%) | 77619.0 → 72535.0 (-6.55%) |
| `p512-k32-owned-repeated` | 12470.0 → 109.0 (-99.13%) | 38.0 → 12.0 (-68.42%) | 947578.0 → 75001.0 (-92.08%) | 951771.0 → 79194.0 (-91.68%) | 3966583.0 → 3961504.0 (-0.13%) | 4041836.0 → 4036757.0 (-0.13%) | 77619.0 → 72535.0 (-6.55%) |
| `p512-k32-owned-batch` | 12470.0 → 109.0 (-99.13%) | 38.0 → 12.0 (-68.42%) | 947578.0 → 75001.0 (-92.08%) | 951771.0 → 79194.0 (-91.68%) | 3966568.0 → 3961489.0 (-0.13%) | 4041821.0 → 4036742.0 (-0.13%) | 77619.0 → 72535.0 (-6.55%) |
| `p512-k32-file-repeated` | 12470.0 → 109.0 (-99.13%) | 38.0 → 12.0 (-68.42%) | 947578.0 → 75001.0 (-92.08%) | 951771.0 → 79194.0 (-91.68%) | 3412231.0 → 3407152.0 (-0.15%) | 4041831.0 → 4036752.0 (-0.13%) | 77619.0 → 72535.0 (-6.55%) |
| `p512-k32-file-batch` | 12470.0 → 109.0 (-99.13%) | 38.0 → 12.0 (-68.42%) | 947578.0 → 75001.0 (-92.08%) | 951771.0 → 79194.0 (-91.68%) | 3412216.0 → 3407137.0 (-0.15%) | 4041816.0 → 4036737.0 (-0.13%) | 77619.0 → 72535.0 (-6.55%) |

## Warnings

No warnings.

The allocation probe's instrumented elapsed time is excluded from native timing comparisons.
