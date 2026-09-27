# 0796 attribute-boundary native and Callgrind analysis

This document is an offline replay of 744 native children and 248 Callgrind children. Native process p50 values use the frozen nearest-rank rule; Callgrind counters are guest diagnostics and do not measure native latency.

## Native paired diagnostics

| Case | Mode | Median after/before p50 | Change | CI low | CI high | Diagnostic regression |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| `distinct-0` | `construct` | 8.678752 | 767.875% | 8.248358 | 11.804714 | yes |
| `distinct-0` | `consume` | 1.070628 | 7.063% | 1.069980 | 1.077228 | yes |
| `distinct-1` | `construct` | 8.414662 | 741.466% | 8.154300 | 9.036273 | yes |
| `distinct-1` | `consume` | 0.964448 | -3.555% | 0.920096 | 0.999746 | no |
| `distinct-16` | `construct` | 9.094537 | 809.454% | 8.122462 | 10.034570 | yes |
| `distinct-16` | `consume` | 1.022298 | 2.230% | 0.995007 | 1.031786 | no |
| `distinct-17` | `construct` | 8.403035 | 740.304% | 8.176223 | 9.065736 | yes |
| `distinct-17` | `consume` | 0.982798 | -1.720% | 0.960817 | 1.021349 | no |
| `distinct-32` | `construct` | 8.369911 | 736.991% | 7.913153 | 8.516863 | yes |
| `distinct-32` | `consume` | 0.996332 | -0.367% | 0.983188 | 0.999079 | no |
| `distinct-33` | `construct` | 8.344570 | 734.457% | 7.855877 | 8.682751 | yes |
| `distinct-33` | `consume` | 0.883829 | -11.617% | 0.866771 | 0.888214 | no |
| `distinct-4` | `construct` | 8.340161 | 734.016% | 7.860182 | 8.422103 | yes |
| `distinct-4` | `consume` | 1.001081 | 0.108% | 0.988482 | 1.068491 | no |
| `distinct-5` | `construct` | 7.980640 | 698.064% | 7.760769 | 9.293882 | yes |
| `distinct-5` | `consume` | 0.872759 | -12.724% | 0.824471 | 1.313560 | no |
| `distinct-64` | `construct` | 7.890472 | 689.047% | 7.773698 | 8.155808 | yes |
| `distinct-64` | `consume` | 0.825757 | -17.424% | 0.808593 | 0.838914 | no |
| `distinct-8` | `construct` | 8.389222 | 738.922% | 7.849139 | 8.983221 | yes |
| `distinct-8` | `consume` | 0.923773 | -7.623% | 0.915049 | 0.938279 | no |
| `distinct-9` | `construct` | 8.165874 | 716.587% | 7.799146 | 8.713820 | yes |
| `distinct-9` | `consume` | 0.944431 | -5.557% | 0.940364 | 0.959241 | no |
| `duplicate-long-quoted-after-1` | `construct` | 8.417872 | 741.787% | 8.360565 | 8.717964 | yes |
| `duplicate-long-quoted-after-1` | `consume` | 24.384922 | 2338.492% | 23.758672 | 24.789628 | yes |
| `duplicate-long-quoted-after-33` | `construct` | 8.316582 | 731.658% | 8.064924 | 8.490886 | yes |
| `duplicate-long-quoted-after-33` | `consume` | 0.922036 | -7.796% | 0.907887 | 0.935156 | no |
| `duplicate-long-unterminated-after-1` | `construct` | 7.903120 | 690.312% | 7.836755 | 9.010118 | yes |
| `duplicate-long-unterminated-after-1` | `consume` | 24.514611 | 2351.461% | 23.497580 | 24.729153 | yes |
| `duplicate-long-unterminated-after-33` | `construct` | 8.680725 | 768.072% | 8.070124 | 9.086161 | yes |
| `duplicate-long-unterminated-after-33` | `consume` | 0.921304 | -7.870% | 0.914000 | 0.928827 | no |
| `duplicate-unquoted-after-1` | `construct` | 8.414615 | 741.462% | 7.653750 | 9.150169 | yes |
| `duplicate-unquoted-after-1` | `consume` | 1.029144 | 2.914% | 1.013976 | 1.059948 | no |
| `duplicate-unquoted-after-33` | `construct` | 8.987418 | 798.742% | 8.437690 | 9.951022 | yes |
| `duplicate-unquoted-after-33` | `consume` | 0.878429 | -12.157% | 0.864540 | 0.942893 | no |
| `duplicate-valid-after-1` | `construct` | 8.132947 | 713.295% | 7.362705 | 8.752609 | yes |
| `duplicate-valid-after-1` | `consume` | 0.979818 | -2.018% | 0.971755 | 1.018329 | no |
| `duplicate-valid-after-32` | `construct` | 8.440135 | 744.013% | 7.927572 | 8.722597 | yes |
| `duplicate-valid-after-32` | `consume` | 0.532182 | -46.782% | 0.521305 | 0.537795 | no |
| `duplicate-valid-after-33` | `construct` | 8.364790 | 736.479% | 8.294350 | 9.628314 | yes |
| `duplicate-valid-after-33` | `consume` | 0.877537 | -12.246% | 0.874286 | 0.881658 | no |
| `duplicate-valid-after-4` | `construct` | 8.384268 | 738.427% | 8.085329 | 9.280340 | yes |
| `duplicate-valid-after-4` | `consume` | 1.008693 | 0.869% | 0.968047 | 1.019314 | no |
| `duplicate-valid-after-5` | `construct` | 8.437969 | 743.797% | 8.331243 | 9.803894 | yes |
| `duplicate-valid-after-5` | `consume` | 0.890841 | -10.916% | 0.817399 | 0.896820 | no |
| `syntax-equals-value-after-0` | `construct` | 8.342820 | 734.282% | 7.861805 | 9.068239 | yes |
| `syntax-equals-value-after-0` | `consume` | 1.311877 | 31.188% | 1.249726 | 1.403088 | yes |
| `syntax-equals-value-after-33` | `construct` | 7.895652 | 689.565% | 7.809898 | 8.341835 | yes |
| `syntax-equals-value-after-33` | `consume` | 0.869197 | -13.080% | 0.861235 | 0.899615 | no |
| `syntax-equals-value-after-4` | `construct` | 8.128788 | 712.879% | 7.847790 | 8.399298 | yes |
| `syntax-equals-value-after-4` | `consume` | 1.090622 | 9.062% | 1.059281 | 1.094502 | yes |
| `syntax-flag-after-0` | `construct` | 8.399832 | 739.983% | 7.884486 | 9.070554 | yes |
| `syntax-flag-after-0` | `consume` | 1.309300 | 30.930% | 1.261347 | 1.370248 | yes |
| `syntax-flag-after-33` | `construct` | 8.398904 | 739.890% | 8.167032 | 8.840641 | yes |
| `syntax-flag-after-33` | `consume` | 0.876618 | -12.338% | 0.869594 | 0.903547 | no |
| `syntax-flag-after-4` | `construct` | 8.390472 | 739.047% | 8.142528 | 9.020236 | yes |
| `syntax-flag-after-4` | `consume` | 1.085214 | 8.521% | 1.031397 | 1.102691 | yes |
| `syntax-unique-tail-after-0` | `construct` | 8.369086 | 736.909% | 8.337352 | 8.709088 | yes |
| `syntax-unique-tail-after-0` | `consume` | 1.065340 | 6.534% | 1.059393 | 1.076657 | yes |
| `syntax-unique-tail-after-33` | `construct` | 8.389840 | 738.984% | 7.993177 | 9.013448 | yes |
| `syntax-unique-tail-after-33` | `consume` | 0.885603 | -11.440% | 0.871227 | 0.900597 | no |
| `syntax-unique-tail-after-4` | `construct` | 8.346543 | 734.654% | 7.903712 | 9.594276 | yes |
| `syntax-unique-tail-after-4` | `consume` | 0.911508 | -8.849% | 0.869567 | 0.938518 | no |

Diagnostic flags: 39. A flag means the median ratio is above 1.05 and the bootstrap lower endpoint is above 1.0; it is not an adoption decision.

Process p50 spread flags above 5%: 66.

## Callgrind profile diagnostics

The parser checked 248 positive dumps and 248 empty termination dumps over the events `Ir, Bc, Bcm, Bi, Bim`.

Exact owner qualification succeeded for 248 profiles and failed for 0. Failures remain in the JSON evidence and are not treated as successful attribution.

The selected self-function census covers checked-iterator, IterState, drop, allocator-name, and lexical-matching symbols. The allocator group is lexical name matching only; it does not count allocator API calls. Inclusive graph costs overlap, so no fractions are reported.

The full receipt identities, raw Callgrind conservation records, owner qualification state, and selected symbol rows are retained in `analysis.json`.
