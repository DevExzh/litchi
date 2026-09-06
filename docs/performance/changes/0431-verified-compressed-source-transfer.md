# 0431: verified compressed source-part transfer

PPTX cross-slide publication previously recompressed copied media. The new ZIP
reader issues an opaque compressed payload only after layout, CRC, decoded
size/content and Deflate termination checks. OPC binds it to the originating
source and execution context. PPTX keeps existing candidate, dependency and
semantic validation; new destination members receive canonical wrappers.
Untouched destination members retain physical preservation.

The [final bundle](../results/change-0431/README.md) retains 16 fresh processes
and 480 samples per role, with forward/reverse repeats and unchanged harness.
API median milliseconds, R1 / R2:

| Media provider | Before | Refined after | Change |
|---|---:|---:|---:|
| Bytes | 252.841 / 257.666 | 28.294 / 28.263 | -88.8% / -89.0% |
| Warm file | 255.177 / 255.196 | 34.008 / 34.393 | -86.7% / -86.5% |
| 4 KiB range | 263.248 / 254.163 | 29.694 / 29.791 | -88.7% / -88.3% |
| 64 KiB, 200 µs range | 654.019 / 660.341 | 527.453 / 527.441 | -19.4% / -20.1% |

Plain API changes are -1.045% to +0.091%. No API median or repeat review flag
exceeds 5%. The [first attempt](../results/change-0431-first-attempt/README.md)
regressed delayed-range latency by 8.8–10%; its full evidence is retained.
Refining capture from 16 KiB to bounded 64 KiB requests removes 768 publication
source calls for the large-range provider with identical reported outputs,
input bytes and work. The provider's delay is simulated, and files are warm.

Both candidates add 16,786,581 source input bytes and 33,559,592 charged work
units relative to baseline. Media process peak RSS stays within 783.2–785.3 MiB;
endpoint RSS has variable entry residency. The writer still buffers generated
members. No causal memory, allocation-count, physical-copy or zero-copy claim
is made. Portable verification rechecks retained source-bound producer reports;
this harness does not export output archives for an independent replay reader.

Implementation commits are `e68c1decb472e4823c7bea9322f12419ddc51bee` and
`0556401e21f1ba740ff033f67c4d4d2151741b8a`. The refined affected suite passes
1,755 tests with five ignored; strict Clippy, formatting and a fresh 1,000-run
ASAN fuzz smoke pass. Earlier workspace/default-feature, rustdoc and ownership
checks are retained, alongside development failures and unchanged unrelated
formatting debt. Source mutation, cancellation, malformed transfer, budgets,
policy refusal and exact partial sink progress have regression coverage.

This is a measured synthetic source-backed publication optimization. Native
OPC image transfer tests do not establish native PPTX cross-slide coverage.
The broader non-iWork goal remains active.
