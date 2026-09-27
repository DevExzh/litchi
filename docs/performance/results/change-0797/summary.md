# 0797 attribute-boundary native and Callgrind analysis

This document is an offline replay of 792 native children and 264 Callgrind children. Native process p50 values use the frozen nearest-rank rule; Callgrind counters are guest diagnostics and do not measure native latency.

## Native paired diagnostics

| Case | Mode | Median after/before p50 | Change | CI low | CI high | Diagnostic regression |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| `distinct-0` | `construct` | 0.923207 | -7.679% | 0.922428 | 1.074916 | no |
| `distinct-0` | `consume` | 1.050633 | 5.063% | 1.050351 | 1.348677 | yes |
| `distinct-1` | `construct` | 0.924115 | -7.589% | 0.923273 | 0.924958 | no |
| `distinct-1` | `consume` | 0.642317 | -35.768% | 0.621102 | 0.643662 | no |
| `distinct-16` | `construct` | 0.922428 | -7.757% | 0.921652 | 0.924894 | no |
| `distinct-16` | `consume` | 1.055608 | 5.561% | 1.047777 | 1.071422 | yes |
| `distinct-17` | `construct` | 0.924115 | -7.589% | 0.923272 | 0.925738 | no |
| `distinct-17` | `consume` | 1.064966 | 6.497% | 1.063425 | 1.096548 | yes |
| `distinct-2` | `construct` | 0.923337 | -7.666% | 0.920960 | 1.000843 | no |
| `distinct-2` | `consume` | 1.429871 | 42.987% | 1.426508 | 1.453885 | yes |
| `distinct-3` | `construct` | 0.923337 | -7.666% | 0.922428 | 0.924958 | no |
| `distinct-3` | `consume` | 1.333872 | 33.387% | 1.310475 | 1.353962 | yes |
| `distinct-32` | `construct` | 0.924895 | -7.510% | 0.922494 | 0.925676 | no |
| `distinct-32` | `consume` | 1.037729 | 3.773% | 1.029105 | 1.045308 | no |
| `distinct-33` | `construct` | 0.924115 | -7.589% | 0.921652 | 0.999938 | no |
| `distinct-33` | `consume` | 1.018424 | 1.842% | 1.016192 | 1.042140 | no |
| `distinct-4` | `construct` | 0.924895 | -7.510% | 0.924051 | 1.076728 | no |
| `distinct-4` | `consume` | 1.270328 | 27.033% | 1.249569 | 1.291744 | yes |
| `distinct-5` | `construct` | 0.923272 | -7.673% | 0.919463 | 1.000781 | no |
| `distinct-5` | `consume` | 1.165349 | 16.535% | 1.154222 | 1.182471 | yes |
| `distinct-64` | `construct` | 0.924958 | -7.504% | 0.922559 | 1.074138 | no |
| `distinct-64` | `consume` | 1.007392 | 0.739% | 0.998386 | 1.019682 | no |
| `distinct-8` | `construct` | 0.923207 | -7.679% | 0.921718 | 0.999936 | no |
| `distinct-8` | `consume` | 1.095853 | 9.585% | 1.086595 | 1.106744 | yes |
| `distinct-9` | `construct` | 0.924037 | -7.596% | 0.922494 | 0.924115 | no |
| `distinct-9` | `consume` | 1.072519 | 7.252% | 1.063560 | 1.078921 | yes |
| `duplicate-long-quoted-after-1` | `construct` | 0.924115 | -7.589% | 0.921784 | 1.000000 | no |
| `duplicate-long-quoted-after-1` | `consume` | 1.366083 | 36.608% | 1.355620 | 1.391794 | yes |
| `duplicate-long-quoted-after-33` | `construct` | 0.922559 | -7.744% | 0.921718 | 0.924895 | no |
| `duplicate-long-quoted-after-33` | `consume` | 1.014198 | 1.420% | 1.008508 | 1.026901 | no |
| `duplicate-long-unterminated-after-1` | `construct` | 0.923337 | -7.666% | 0.920875 | 0.999872 | no |
| `duplicate-long-unterminated-after-1` | `consume` | 1.367566 | 36.757% | 1.356111 | 1.400956 | yes |
| `duplicate-long-unterminated-after-33` | `construct` | 0.922559 | -7.744% | 0.922428 | 0.924115 | no |
| `duplicate-long-unterminated-after-33` | `consume` | 1.015152 | 1.515% | 1.009632 | 1.031313 | no |
| `duplicate-unquoted-after-1` | `construct` | 0.924115 | -7.589% | 0.923272 | 0.924179 | no |
| `duplicate-unquoted-after-1` | `consume` | 1.368122 | 36.812% | 1.353461 | 1.423886 | yes |
| `duplicate-unquoted-after-33` | `construct` | 0.925676 | -7.432% | 0.922494 | 1.001622 | no |
| `duplicate-unquoted-after-33` | `consume` | 1.017377 | 1.738% | 1.014473 | 1.031731 | no |
| `duplicate-valid-after-1` | `construct` | 0.924115 | -7.589% | 0.922494 | 0.924115 | no |
| `duplicate-valid-after-1` | `consume` | 1.389820 | 38.982% | 1.356750 | 1.412975 | yes |
| `duplicate-valid-after-32` | `construct` | 0.922428 | -7.757% | 0.920102 | 0.999286 | no |
| `duplicate-valid-after-32` | `consume` | 1.022576 | 2.258% | 1.013287 | 1.024805 | no |
| `duplicate-valid-after-33` | `construct` | 0.924115 | -7.589% | 0.922428 | 0.924115 | no |
| `duplicate-valid-after-33` | `consume` | 1.025709 | 2.571% | 1.020370 | 1.040848 | no |
| `duplicate-valid-after-4` | `construct` | 0.924115 | -7.589% | 0.923272 | 0.924964 | no |
| `duplicate-valid-after-4` | `consume` | 1.251565 | 25.157% | 1.244676 | 1.306002 | yes |
| `duplicate-valid-after-5` | `construct` | 0.924115 | -7.589% | 0.922494 | 0.924115 | no |
| `duplicate-valid-after-5` | `consume` | 1.212622 | 21.262% | 1.187323 | 1.229298 | yes |
| `syntax-equals-value-after-0` | `construct` | 0.924115 | -7.589% | 0.922494 | 0.924958 | no |
| `syntax-equals-value-after-0` | `consume` | 1.051323 | 5.132% | 1.020014 | 1.069879 | yes |
| `syntax-equals-value-after-33` | `construct` | 0.923272 | -7.673% | 0.921652 | 1.000843 | no |
| `syntax-equals-value-after-33` | `consume` | 1.018726 | 1.873% | 1.009558 | 1.027404 | no |
| `syntax-equals-value-after-4` | `construct` | 0.924115 | -7.589% | 0.923272 | 0.925738 | no |
| `syntax-equals-value-after-4` | `consume` | 1.240537 | 24.054% | 1.227286 | 1.261519 | yes |
| `syntax-flag-after-0` | `construct` | 0.922428 | -7.757% | 0.921706 | 0.923272 | no |
| `syntax-flag-after-0` | `consume` | 1.051848 | 5.185% | 1.023870 | 1.075628 | yes |
| `syntax-flag-after-33` | `construct` | 0.923273 | -7.673% | 0.921718 | 0.999241 | no |
| `syntax-flag-after-33` | `consume` | 1.021283 | 2.128% | 1.018654 | 1.037844 | no |
| `syntax-flag-after-4` | `construct` | 0.923337 | -7.666% | 0.922494 | 0.924895 | no |
| `syntax-flag-after-4` | `consume` | 1.252198 | 25.220% | 1.234222 | 1.264327 | yes |
| `syntax-unique-tail-after-0` | `construct` | 0.923337 | -7.666% | 0.921652 | 1.001546 | no |
| `syntax-unique-tail-after-0` | `consume` | 0.800047 | -19.995% | 0.795135 | 0.824381 | no |
| `syntax-unique-tail-after-33` | `construct` | 0.923337 | -7.666% | 0.922494 | 0.999094 | no |
| `syntax-unique-tail-after-33` | `consume` | 1.013221 | 1.322% | 0.999181 | 1.019403 | no |
| `syntax-unique-tail-after-4` | `construct` | 0.923337 | -7.666% | 0.921652 | 1.000000 | no |
| `syntax-unique-tail-after-4` | `consume` | 1.196146 | 19.615% | 1.171540 | 1.210534 | yes |

Diagnostic flags: 20. A flag means the median ratio is above 1.05 and the bootstrap lower endpoint is above 1.0; it is not an adoption decision.

Process p50 spread flags above 5%: 22.

## Callgrind profile diagnostics

The parser checked 264 positive dumps and 264 empty termination dumps over the events `Ir, Bc, Bcm, Bi, Bim`.

Exact owner qualification succeeded for 264 profiles and failed for 0. Failures remain in the JSON evidence and are not treated as successful attribution.

The selected self-function census covers checked-iterator, IterState, drop, allocator-name, and lexical-matching symbols. The allocator group is lexical name matching only; it does not count allocator API calls. Inclusive graph costs overlap, so no fractions are reported.

The full receipt identities, raw Callgrind conservation records, owner qualification state, and selected symbol rows are retained in `analysis.json`.
