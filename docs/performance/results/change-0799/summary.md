# 0799 attribute-boundary native and Callgrind analysis

This document is an offline replay of 936 native children and 312 Callgrind children. Native process p50 values use the frozen nearest-rank rule; Callgrind counters are guest diagnostics and do not measure native latency.

## Native paired diagnostics

| Case | Mode | Median after/before p50 | Change | CI low | CI high | Diagnostic regression |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| `distinct-0` | `construct` | 0.922428 | -7.757% | 0.889121 | 0.923272 | no |
| `distinct-0` | `consume` | 0.628882 | -37.112% | 0.624174 | 1.050778 | no |
| `distinct-1` | `construct` | 0.922428 | -7.757% | 0.920942 | 0.923356 | no |
| `distinct-1` | `consume` | 0.651018 | -34.898% | 0.627111 | 0.655196 | no |
| `distinct-16` | `construct` | 0.923337 | -7.666% | 0.920864 | 0.924115 | no |
| `distinct-16` | `consume` | 1.117225 | 11.722% | 1.087154 | 1.144602 | yes |
| `distinct-17` | `construct` | 0.924115 | -7.589% | 0.922494 | 0.924958 | no |
| `distinct-17` | `consume` | 1.112293 | 11.229% | 1.096143 | 1.123549 | yes |
| `distinct-2` | `construct` | 0.924037 | -7.596% | 0.922494 | 0.924832 | no |
| `distinct-2` | `consume` | 0.832837 | -16.716% | 0.832013 | 0.838709 | no |
| `distinct-3` | `construct` | 0.921652 | -7.835% | 0.889838 | 0.924115 | no |
| `distinct-3` | `consume` | 1.552761 | 55.276% | 1.535245 | 1.584619 | yes |
| `distinct-32` | `construct` | 0.923207 | -7.679% | 0.890681 | 0.924831 | no |
| `distinct-32` | `consume` | 1.060278 | 6.028% | 1.055744 | 1.071060 | yes |
| `distinct-33` | `construct` | 0.923273 | -7.673% | 0.889971 | 0.924115 | no |
| `distinct-33` | `consume` | 1.027661 | 2.766% | 1.010677 | 1.033208 | no |
| `distinct-4` | `construct` | 0.924115 | -7.589% | 0.921652 | 0.924895 | no |
| `distinct-4` | `consume` | 1.431872 | 43.187% | 1.411620 | 1.487191 | yes |
| `distinct-5` | `construct` | 0.924115 | -7.589% | 0.922428 | 0.924115 | no |
| `distinct-5` | `consume` | 1.246714 | 24.671% | 1.227802 | 1.258834 | yes |
| `distinct-64` | `construct` | 0.923273 | -7.673% | 0.921669 | 0.924115 | no |
| `distinct-64` | `consume` | 1.022225 | 2.222% | 1.017452 | 1.031382 | no |
| `distinct-8` | `construct` | 0.922494 | -7.751% | 0.921736 | 0.924958 | no |
| `distinct-8` | `consume` | 1.198661 | 19.866% | 1.156878 | 1.221393 | yes |
| `distinct-9` | `construct` | 0.924115 | -7.589% | 0.923337 | 0.924115 | no |
| `distinct-9` | `consume` | 1.152628 | 15.263% | 1.126957 | 1.166380 | yes |
| `duplicate-long-quoted-after-1` | `construct` | 0.924115 | -7.589% | 0.921652 | 0.924179 | no |
| `duplicate-long-quoted-after-1` | `consume` | 0.641187 | -35.881% | 0.630486 | 0.652401 | no |
| `duplicate-long-quoted-after-2` | `construct` | 0.922428 | -7.757% | 0.921641 | 0.924817 | no |
| `duplicate-long-quoted-after-2` | `consume` | 1.508976 | 50.898% | 1.488662 | 1.518779 | yes |
| `duplicate-long-quoted-after-33` | `construct` | 0.923338 | -7.666% | 0.921650 | 0.926503 | no |
| `duplicate-long-quoted-after-33` | `consume` | 1.018904 | 1.890% | 1.010579 | 1.030027 | no |
| `duplicate-long-unterminated-after-1` | `construct` | 0.923273 | -7.673% | 0.922494 | 0.924958 | no |
| `duplicate-long-unterminated-after-1` | `consume` | 0.649098 | -35.090% | 0.600610 | 0.656932 | no |
| `duplicate-long-unterminated-after-2` | `construct` | 0.921652 | -7.835% | 0.889905 | 0.924052 | no |
| `duplicate-long-unterminated-after-2` | `consume` | 1.535865 | 53.586% | 1.503638 | 1.589069 | yes |
| `duplicate-long-unterminated-after-33` | `construct` | 0.921652 | -7.835% | 0.889166 | 0.924959 | no |
| `duplicate-long-unterminated-after-33` | `consume` | 1.018662 | 1.866% | 1.015437 | 1.021936 | no |
| `duplicate-unquoted-after-1` | `construct` | 0.923356 | -7.664% | 0.921718 | 0.924895 | no |
| `duplicate-unquoted-after-1` | `consume` | 0.648781 | -35.122% | 0.640839 | 0.658460 | no |
| `duplicate-unquoted-after-33` | `construct` | 0.924115 | -7.589% | 0.922363 | 0.924895 | no |
| `duplicate-unquoted-after-33` | `consume` | 1.032702 | 3.270% | 1.024370 | 1.036415 | no |
| `duplicate-valid-after-1` | `construct` | 0.924051 | -7.595% | 0.922559 | 0.925738 | no |
| `duplicate-valid-after-1` | `consume` | 0.636832 | -36.317% | 0.633056 | 0.642358 | no |
| `duplicate-valid-after-2` | `construct` | 0.923273 | -7.673% | 0.921652 | 0.924895 | no |
| `duplicate-valid-after-2` | `consume` | 1.518736 | 51.874% | 1.494876 | 1.566236 | yes |
| `duplicate-valid-after-32` | `construct` | 0.922494 | -7.751% | 0.922428 | 0.924115 | no |
| `duplicate-valid-after-32` | `consume` | 1.041423 | 4.142% | 1.029080 | 1.071784 | no |
| `duplicate-valid-after-33` | `construct` | 0.924115 | -7.589% | 0.923129 | 0.925738 | no |
| `duplicate-valid-after-33` | `consume` | 1.026491 | 2.649% | 1.020055 | 1.049278 | no |
| `duplicate-valid-after-4` | `construct` | 0.923259 | -7.674% | 0.920038 | 0.924895 | no |
| `duplicate-valid-after-4` | `consume` | 1.393453 | 39.345% | 1.362262 | 1.425057 | yes |
| `duplicate-valid-after-5` | `construct` | 0.923272 | -7.673% | 0.920809 | 0.925598 | no |
| `duplicate-valid-after-5` | `consume` | 1.249514 | 24.951% | 1.224106 | 1.272289 | yes |
| `syntax-equals-value-after-0` | `construct` | 0.923273 | -7.673% | 0.889898 | 0.924115 | no |
| `syntax-equals-value-after-0` | `consume` | 1.024949 | 2.495% | 0.984572 | 1.065757 | no |
| `syntax-equals-value-after-2` | `construct` | 0.922428 | -7.757% | 0.890616 | 0.924115 | no |
| `syntax-equals-value-after-2` | `consume` | 1.559910 | 55.991% | 1.529931 | 1.571839 | yes |
| `syntax-equals-value-after-33` | `construct` | 0.922513 | -7.749% | 0.920100 | 0.924115 | no |
| `syntax-equals-value-after-33` | `consume` | 1.034171 | 3.417% | 1.025951 | 1.053997 | no |
| `syntax-equals-value-after-4` | `construct` | 0.923207 | -7.679% | 0.922428 | 0.925740 | no |
| `syntax-equals-value-after-4` | `consume` | 1.405048 | 40.505% | 1.395458 | 1.408185 | yes |
| `syntax-flag-after-0` | `construct` | 0.921652 | -7.835% | 0.888347 | 0.924115 | no |
| `syntax-flag-after-0` | `consume` | 1.029861 | 2.986% | 1.011335 | 1.067366 | no |
| `syntax-flag-after-2` | `construct` | 0.924051 | -7.595% | 0.890681 | 0.924958 | no |
| `syntax-flag-after-2` | `consume` | 1.557298 | 55.730% | 1.518724 | 1.573107 | yes |
| `syntax-flag-after-33` | `construct` | 0.924051 | -7.595% | 0.921650 | 0.925801 | no |
| `syntax-flag-after-33` | `consume` | 1.038014 | 3.801% | 1.033561 | 1.046833 | no |
| `syntax-flag-after-4` | `construct` | 0.922559 | -7.744% | 0.922494 | 0.924978 | no |
| `syntax-flag-after-4` | `consume` | 1.389484 | 38.948% | 1.383389 | 1.419355 | yes |
| `syntax-unique-tail-after-0` | `construct` | 0.924959 | -7.504% | 0.922494 | 0.926583 | no |
| `syntax-unique-tail-after-0` | `consume` | 0.796441 | -20.356% | 0.780762 | 0.831312 | no |
| `syntax-unique-tail-after-2` | `construct` | 0.923207 | -7.679% | 0.920875 | 0.925676 | no |
| `syntax-unique-tail-after-2` | `consume` | 1.503305 | 50.330% | 1.498072 | 1.525634 | yes |
| `syntax-unique-tail-after-33` | `construct` | 0.922494 | -7.751% | 0.922428 | 0.924115 | no |
| `syntax-unique-tail-after-33` | `consume` | 1.024937 | 2.494% | 1.018563 | 1.041208 | no |
| `syntax-unique-tail-after-4` | `construct` | 0.923272 | -7.673% | 0.920875 | 0.924895 | no |
| `syntax-unique-tail-after-4` | `consume` | 1.318737 | 31.874% | 1.308147 | 1.325316 | yes |

Diagnostic flags: 19. A flag means the median ratio is above 1.05 and the bootstrap lower endpoint is above 1.0; it is not an adoption decision.

Process p50 spread flags above 5%: 23.

## Callgrind profile diagnostics

The parser checked 312 positive dumps and 312 empty termination dumps over the events `Ir, Bc, Bcm, Bi, Bim`.

Exact owner qualification succeeded for 312 profiles and failed for 0. Failures remain in the JSON evidence and are not treated as successful attribution.

The selected self-function census covers checked-iterator, IterState, drop, allocator-name, and lexical-matching symbols. The allocator group is lexical name matching only; it does not count allocator API calls. Inclusive graph costs overlap, so no fractions are reported.

The full receipt identities, raw Callgrind conservation records, owner qualification state, and selected symbol rows are retained in `analysis.json`.
