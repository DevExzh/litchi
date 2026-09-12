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
| `p128-k1-owned-repeated` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3815605.0 → 3815610.0 (+0.00%) | 3852590.0 → 3852595.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p128-k1-owned-batch` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3815590.0 → 3815595.0 (+0.00%) | 3852575.0 → 3852580.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p128-k1-file-repeated` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3282757.0 → 3282762.0 (+0.00%) | 3832742.0 → 3832747.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p128-k1-file-batch` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3282742.0 → 3282747.0 (+0.00%) | 3832727.0 → 3832732.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p128-k8-owned-repeated` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3818286.0 → 3818291.0 (+0.00%) | 3853248.0 → 3853253.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p128-k8-owned-batch` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3818271.0 → 3818276.0 (+0.00%) | 3853233.0 → 3853238.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p128-k8-file-repeated` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3285438.0 → 3285443.0 (+0.00%) | 3837446.0 → 3837451.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p128-k8-file-batch` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3285423.0 → 3285428.0 (+0.00%) | 3837431.0 → 3837436.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p128-k32-owned-repeated` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3827483.0 → 3827488.0 (+0.00%) | 3855509.0 → 3855514.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p128-k32-owned-batch` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3827468.0 → 3827473.0 (+0.00%) | 3855494.0 → 3855499.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p128-k32-file-repeated` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3294635.0 → 3294640.0 (+0.00%) | 3853579.0 → 3853584.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p128-k32-file-batch` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3294620.0 → 3294625.0 (+0.00%) | 3853564.0 → 3853569.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p512-k1-owned-repeated` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3949621.0 → 3949626.0 (+0.00%) | 4015915.0 → 4015920.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p512-k1-owned-batch` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3949606.0 → 3949611.0 (+0.00%) | 4015900.0 → 4015905.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p512-k1-file-repeated` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3395269.0 → 3395274.0 (+0.00%) | 4015910.0 → 4015915.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p512-k1-file-batch` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3395254.0 → 3395259.0 (+0.00%) | 4015895.0 → 4015900.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p512-k8-owned-repeated` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3952302.0 → 3952307.0 (+0.00%) | 4020619.0 → 4020624.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p512-k8-owned-batch` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3952287.0 → 3952292.0 (+0.00%) | 4020604.0 → 4020609.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p512-k8-file-repeated` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3397950.0 → 3397955.0 (+0.00%) | 4020614.0 → 4020619.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p512-k8-file-batch` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3397935.0 → 3397940.0 (+0.00%) | 4020599.0 → 4020604.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p512-k32-owned-repeated` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3961499.0 → 3961504.0 (+0.00%) | 4036752.0 → 4036757.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p512-k32-owned-batch` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3961484.0 → 3961489.0 (+0.00%) | 4036737.0 → 4036742.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p512-k32-file-repeated` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3407147.0 → 3407152.0 (+0.00%) | 4036747.0 → 4036752.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |
| `p512-k32-file-batch` | 109.0 → 96.0 (-11.93%) | 12.0 → 6.0 (-50.00%) | 75001.0 → 74218.0 (-1.04%) | 79194.0 → 78411.0 (-0.99%) | 3407132.0 → 3407137.0 (+0.00%) | 4036732.0 → 4036737.0 (+0.00%) | 72535.0 → 72535.0 (+0.00%) |

## Warnings

No warnings.

The allocation probe's instrumented elapsed time is excluded from native timing comparisons.
