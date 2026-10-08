# 0972 — offline recovery of the 0971 callchain decode

**Status: diagnostic recovery complete; the 0971 native gate remains failed.**
This record corrects the 0971 offline address binding using the sealed child-01
capture. It performed no new capture, target build, Office workload, production
edit, or mutation of the retained `perf.data`.

## Address and raw-scan correction

The 0971 scanner used target load bias `0x64d1e4e5c000`; the verified
file-offset-zero ELF load bias is `0x64d1e4e5b000`. It also used a truncated
executable interval. The corrected scanner derives the interval from the
retained ELF `PF_X PT_LOAD` range and scans all 1,387 sample records. Its
potential address-word matches are 53 marker-range, 39 marker-return, 17
`measure_one`-return, and 54 `run_case`-range samples in the 908-record
target-leaf subset. These are stack-word matches only: they do not prove live
ancestry, phase ownership, stack depth, or performance. The earlier 0971
raw-zero conclusion is withdrawn.

## Authoritative offline replay

The sealed v6 replay uses an external exact-DSO diagnostic interposer. It
forwards the complete `dwfl_report_elf` ABI, supplies the verified target base,
and adjusts the returned module comparator start only for direct `/usr/bin/perf`
callers. Non-target modules and non-perf callers are unchanged; stock and no-op
controls are byte-identical. The original input hash remains
`499e3008e6496b01fe09c9cf7f6f2aedbf46c1b24f7697e98289c09460143746`.

| Decode | Sample records | Empty chains | Frames | Marker chains |
| --- | ---: | ---: | ---: | ---: |
| stock | 1,387 | 906 | 1,945 | 0 |
| no-op interposer | 1,387 | 906 | 1,945 | 0 |
| v6 corrected | 1,387 | 0 | 11,154 | 78 |

The corrected decode leaves 611 unresolved frame lines, all 78 marker chains
have resolved marker frames, and 41 include the outer `measure_one` frame.
Thirty-seven marker chains lack that outer measurement context. The marker
chains carry 390,000,000 instructions of inclusive event period; overlapping
inclusive tags are not phase fractions or timing measurements.

## Disposition

The original pilot required at least 100 marker samples. The corrected replay
recovers 78, so the 0971 pilot remains below its gate and the planned 12-child
cohort is not revived. The offline evidence supports a map-metadata/decoder
interaction in this replay, but does not identify the upstream perf/libdw
layer, prove a stack-depth or unwind cause, or establish a live performance
owner. There is no phase attribution, latency result, speedup claim, or
production adoption.

The sealed 0971 packet remains historical evidence of the original native pilot;
only its raw-zero interpretation is corrected. The external diagnostic shim and
all replay work remain outside production. Cleanup preserved the retained
binary, source, raw input, and final controls.

## Evidence

| Evidence | SHA-256 |
| --- | --- |
| [final diagnosis](results/change-0972/analysis/diagnosis-v3.md) | `80c54eeb6116b585b9088399c14eef2583b73202faa4a364f0983968753eff9c` |
| [corrected raw scan v2](results/change-0972/analysis/corrected-raw-stack-scan-0972-v2.json) | `6ef76bc01cc7e298c924eda95cfd29dc0cc070ac3e3de4cbe9d87a7cf1b93a9b` |
| [corrected decode summary](results/change-0972/analysis/dwfl-corrected-decode-v6.json) | `3b75fdabf79ca7104f24e5deb2e389ecbd94c7a439579cc16a5a8d3c7b42e7db` |
| [v6 replay author seal](results/change-0972/receipts/dwfl-replay-author-seal-v6.json) | `a8be68fb63e7460d839ca23ab2457b2415762da76d8a92d26e8640281a7bde1c` |
| [independent offline review](results/change-0972/receipts/independent-offline-review-0972-final-v1.json) | `617900bc7f758eaf1e3c38643538a74dd04e75f6a679c3beea39942b65952c3d` |
| [root binding resolution](results/change-0972/receipts/root-binding-resolution-0972-v1.json) | `c49b9a130e58f3a474b3eae44852400b6a6555a51e3c4d687980be7f2a2a914c` |
| [temporary cleanup](results/change-0972/receipts/temporary-cleanup-0972-v1.json) | `952105c102be9e9cf8313f61acc6b0cb5c30ec37d839ebe72ca7b52b5876b236` |

**Performance claim:** none. **Production change:** none. **Attribution
admission:** failed at the original 100-marker gate.
