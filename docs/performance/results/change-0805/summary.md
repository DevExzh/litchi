# 0805 attribute-boundary production preflight analysis

This document is an offline replay of 936 native children and 312 Callgrind children. Native process p50 values use the frozen nearest-rank rule; Callgrind counters are guest diagnostics and do not measure native latency.

## Native paired diagnostics

| Case | Mode | Median after/before p50 | Change | CI low | CI high | Diagnostic regression |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| `distinct-0` | `construct` | 1.067468 | 6.747% | 0.999158 | 1.163444 | no |
| `distinct-0` | `consume` | 0.318976 | -68.102% | 0.318807 | 0.336264 | no |
| `distinct-1` | `construct` | 1.000000 | 0.000% | 1.000000 | 1.000843 | no |
| `distinct-1` | `consume` | 0.613544 | -38.646% | 0.588405 | 0.617521 | no |
| `distinct-16` | `construct` | 1.000000 | 0.000% | 0.997475 | 1.000000 | no |
| `distinct-16` | `consume` | 0.988219 | -1.178% | 0.984651 | 1.008824 | no |
| `distinct-17` | `construct` | 1.136238 | 13.624% | 1.036975 | 1.159219 | yes |
| `distinct-17` | `consume` | 0.938641 | -6.136% | 0.934897 | 0.953873 | no |
| `distinct-2` | `construct` | 1.002530 | 0.253% | 0.998316 | 1.128162 | no |
| `distinct-2` | `consume` | 0.809200 | -19.080% | 0.791095 | 0.809967 | no |
| `distinct-3` | `construct` | 1.129848 | 12.985% | 1.000927 | 1.164418 | yes |
| `distinct-3` | `consume` | 0.936681 | -6.332% | 0.917889 | 0.949648 | no |
| `distinct-32` | `construct` | 0.999160 | -0.084% | 0.997475 | 1.066526 | no |
| `distinct-32` | `consume` | 1.032739 | 3.274% | 1.012950 | 1.054385 | no |
| `distinct-33` | `construct` | 1.000845 | 0.084% | 0.999158 | 1.135596 | no |
| `distinct-33` | `consume` | 0.556748 | -44.325% | 0.537645 | 0.572747 | no |
| `distinct-4` | `construct` | 1.000000 | 0.000% | 0.998316 | 1.002532 | no |
| `distinct-4` | `consume` | 1.226491 | 22.649% | 1.198300 | 1.243788 | yes |
| `distinct-5` | `construct` | 1.000926 | 0.093% | 0.999160 | 1.074199 | no |
| `distinct-5` | `consume` | 0.969531 | -3.047% | 0.943129 | 0.990174 | no |
| `distinct-64` | `construct` | 1.000000 | 0.000% | 0.999157 | 1.001686 | no |
| `distinct-64` | `consume` | 0.989845 | -1.015% | 0.985159 | 1.011261 | no |
| `distinct-8` | `construct` | 1.000842 | 0.084% | 0.998318 | 1.067456 | no |
| `distinct-8` | `consume` | 1.027703 | 2.770% | 1.023107 | 1.039633 | no |
| `distinct-9` | `construct` | 1.000843 | 0.084% | 0.998316 | 1.152120 | no |
| `distinct-9` | `consume` | 0.920467 | -7.953% | 0.910679 | 0.930129 | no |
| `duplicate-long-quoted-after-1` | `construct` | 1.000000 | 0.000% | 0.998990 | 1.131425 | no |
| `duplicate-long-quoted-after-1` | `consume` | 0.650978 | -34.902% | 0.646127 | 0.658537 | no |
| `duplicate-long-quoted-after-2` | `construct` | 1.058179 | 5.818% | 1.000000 | 1.140809 | no |
| `duplicate-long-quoted-after-2` | `consume` | 0.818288 | -18.171% | 0.809507 | 0.836903 | no |
| `duplicate-long-quoted-after-33` | `construct` | 1.001686 | 0.169% | 1.000842 | 1.070826 | no |
| `duplicate-long-quoted-after-33` | `consume` | 0.570848 | -42.915% | 0.564937 | 0.583826 | no |
| `duplicate-long-unterminated-after-1` | `construct` | 1.000000 | 0.000% | 1.000000 | 1.073356 | no |
| `duplicate-long-unterminated-after-1` | `consume` | 0.653726 | -34.627% | 0.643539 | 0.662836 | no |
| `duplicate-long-unterminated-after-2` | `construct` | 1.000843 | 0.084% | 1.000000 | 1.002530 | no |
| `duplicate-long-unterminated-after-2` | `consume` | 0.827169 | -17.283% | 0.817101 | 0.843378 | no |
| `duplicate-long-unterminated-after-33` | `construct` | 1.002530 | 0.253% | 0.998319 | 1.141560 | no |
| `duplicate-long-unterminated-after-33` | `consume` | 0.562098 | -43.790% | 0.558405 | 0.571510 | no |
| `duplicate-unquoted-after-1` | `construct` | 1.000000 | 0.000% | 0.999157 | 1.143221 | no |
| `duplicate-unquoted-after-1` | `consume` | 0.651778 | -34.822% | 0.643696 | 0.664405 | no |
| `duplicate-unquoted-after-33` | `construct` | 1.000000 | 0.000% | 0.997478 | 1.129633 | no |
| `duplicate-unquoted-after-33` | `consume` | 0.856684 | -14.332% | 0.841151 | 0.868477 | no |
| `duplicate-valid-after-1` | `construct` | 1.000000 | 0.000% | 0.998316 | 1.069983 | no |
| `duplicate-valid-after-1` | `consume` | 0.657779 | -34.222% | 0.647442 | 0.668070 | no |
| `duplicate-valid-after-2` | `construct` | 1.000843 | 0.084% | 0.999158 | 1.091906 | no |
| `duplicate-valid-after-2` | `consume` | 0.827697 | -17.230% | 0.812783 | 0.842014 | no |
| `duplicate-valid-after-32` | `construct` | 1.000843 | 0.084% | 1.000000 | 1.135750 | no |
| `duplicate-valid-after-32` | `consume` | 0.545924 | -45.408% | 0.541914 | 0.578158 | no |
| `duplicate-valid-after-33` | `construct` | 1.000843 | 0.084% | 1.000000 | 1.134689 | no |
| `duplicate-valid-after-33` | `consume` | 0.862169 | -13.783% | 0.855428 | 0.870532 | no |
| `duplicate-valid-after-4` | `construct` | 1.135616 | 13.562% | 0.997489 | 1.189713 | no |
| `duplicate-valid-after-4` | `consume` | 1.107260 | 10.726% | 1.095358 | 1.119992 | yes |
| `duplicate-valid-after-5` | `construct` | 1.000000 | 0.000% | 0.997479 | 1.070911 | no |
| `duplicate-valid-after-5` | `consume` | 0.926661 | -7.334% | 0.879330 | 0.936566 | no |
| `syntax-equals-value-after-0` | `construct` | 1.000001 | 0.000% | 0.998316 | 1.067454 | no |
| `syntax-equals-value-after-0` | `consume` | 1.022081 | 2.208% | 0.967258 | 1.064163 | no |
| `syntax-equals-value-after-2` | `construct` | 1.001689 | 0.169% | 0.999158 | 1.170320 | no |
| `syntax-equals-value-after-2` | `consume` | 0.964001 | -3.600% | 0.945747 | 0.978961 | no |
| `syntax-equals-value-after-33` | `construct` | 1.061551 | 6.155% | 0.998319 | 1.142496 | no |
| `syntax-equals-value-after-33` | `consume` | 0.862570 | -13.743% | 0.845937 | 0.870961 | no |
| `syntax-equals-value-after-4` | `construct` | 1.001686 | 0.169% | 1.000759 | 1.068298 | no |
| `syntax-equals-value-after-4` | `consume` | 1.286817 | 28.682% | 1.280224 | 1.297670 | yes |
| `syntax-flag-after-0` | `construct` | 1.000842 | 0.084% | 0.999158 | 1.168634 | no |
| `syntax-flag-after-0` | `consume` | 1.029831 | 2.983% | 1.003291 | 1.060821 | no |
| `syntax-flag-after-2` | `construct` | 1.000845 | 0.084% | 0.998316 | 1.135870 | no |
| `syntax-flag-after-2` | `consume` | 0.945928 | -5.407% | 0.943412 | 0.963100 | no |
| `syntax-flag-after-33` | `construct` | 1.001689 | 0.169% | 0.999158 | 1.144182 | no |
| `syntax-flag-after-33` | `consume` | 0.861445 | -13.856% | 0.855143 | 0.869222 | no |
| `syntax-flag-after-4` | `construct` | 1.067454 | 6.745% | 1.000000 | 1.144900 | no |
| `syntax-flag-after-4` | `consume` | 1.286344 | 28.634% | 1.209093 | 1.299985 | yes |
| `syntax-unique-tail-after-0` | `construct` | 1.000843 | 0.084% | 0.998319 | 1.158647 | no |
| `syntax-unique-tail-after-0` | `consume` | 0.803487 | -19.651% | 0.777962 | 0.838107 | no |
| `syntax-unique-tail-after-2` | `construct` | 1.000843 | 0.084% | 0.999158 | 1.073234 | no |
| `syntax-unique-tail-after-2` | `consume` | 0.976483 | -2.352% | 0.963244 | 0.989898 | no |
| `syntax-unique-tail-after-33` | `construct` | 1.000000 | 0.000% | 0.998319 | 1.080944 | no |
| `syntax-unique-tail-after-33` | `consume` | 0.882772 | -11.723% | 0.861403 | 0.899036 | no |
| `syntax-unique-tail-after-4` | `construct` | 1.000000 | 0.000% | 0.999158 | 1.001686 | no |
| `syntax-unique-tail-after-4` | `consume` | 0.996584 | -0.342% | 0.980218 | 1.003733 | no |

Diagnostic flags: 6. A flag means the median ratio is above 1.05 and the bootstrap lower endpoint is above 1.0; it is not an adoption decision.

Process p50 spread flags above 5%: 41.

## Callgrind profile diagnostics

The parser checked 312 positive dumps and 312 empty termination dumps over the events `Ir, Bc, Bcm, Bi, Bim`.

Exact owner qualification succeeded for 312 profiles and failed for 0. Failures remain in the JSON evidence and are not treated as successful attribution.

The selected self-function census covers checked-iterator, IterState, drop, allocator-name, and lexical-matching symbols. The allocator group is lexical name matching only; it does not count allocator API calls. Inclusive graph costs overlap, so no fractions are reported.

The full receipt identities, raw Callgrind conservation records, owner qualification state, and selected symbol rows are retained in `analysis.json`.
