# 0804 attribute-boundary native and Callgrind analysis

This document is an offline replay of 936 native children and 312 Callgrind children. Native process p50 values use the frozen nearest-rank rule; Callgrind counters are guest diagnostics and do not measure native latency.

## Native paired diagnostics

| Case | Mode | Median after/before p50 | Change | CI low | CI high | Diagnostic regression |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| `distinct-0` | `construct` | 1.014416 | 1.442% | 0.607013 | 1.137437 | no |
| `distinct-0` | `consume` | 1.012988 | 1.299% | 0.997856 | 1.067949 | no |
| `distinct-1` | `construct` | 1.000000 | 0.000% | 0.999157 | 1.155177 | no |
| `distinct-1` | `consume` | 1.002094 | 0.209% | 1.001698 | 1.002979 | no |
| `distinct-16` | `construct` | 0.999158 | -0.084% | 0.936474 | 1.120480 | no |
| `distinct-16` | `consume` | 0.844107 | -15.589% | 0.812481 | 0.864115 | no |
| `distinct-17` | `construct` | 0.944528 | -5.547% | 0.873445 | 1.001686 | no |
| `distinct-17` | `consume` | 0.812722 | -18.728% | 0.790919 | 0.826626 | no |
| `distinct-2` | `construct` | 1.000843 | 0.084% | 0.939911 | 1.001686 | no |
| `distinct-2` | `consume` | 1.012556 | 1.256% | 1.009678 | 1.014030 | no |
| `distinct-3` | `construct` | 0.996282 | -0.372% | 0.922598 | 1.000000 | no |
| `distinct-3` | `consume` | 0.999387 | -0.061% | 0.995718 | 1.005889 | no |
| `distinct-32` | `construct` | 1.128906 | 12.891% | 0.995643 | 1.155884 | no |
| `distinct-32` | `consume` | 0.751753 | -24.825% | 0.725271 | 0.761829 | no |
| `distinct-33` | `construct` | 0.976303 | -2.370% | 0.863291 | 1.128061 | no |
| `distinct-33` | `consume` | 0.757407 | -24.259% | 0.741482 | 0.765053 | no |
| `distinct-4` | `construct` | 0.997475 | -0.253% | 0.866449 | 1.080946 | no |
| `distinct-4` | `consume` | 1.010841 | 1.084% | 1.008342 | 1.017765 | no |
| `distinct-5` | `construct` | 1.001686 | 0.169% | 0.998319 | 1.142496 | no |
| `distinct-5` | `consume` | 0.996842 | -0.316% | 0.987939 | 1.017724 | no |
| `distinct-64` | `construct` | 0.998315 | -0.168% | 0.861088 | 1.070669 | no |
| `distinct-64` | `consume` | 0.908234 | -9.177% | 0.877709 | 0.920927 | no |
| `distinct-8` | `construct` | 0.989737 | -1.026% | 0.878298 | 1.066610 | no |
| `distinct-8` | `consume` | 0.981138 | -1.886% | 0.958503 | 0.996608 | no |
| `distinct-9` | `construct` | 1.000000 | 0.000% | 0.941037 | 1.133960 | no |
| `distinct-9` | `consume` | 0.971141 | -2.886% | 0.961567 | 0.991884 | no |
| `duplicate-long-quoted-after-1` | `construct` | 1.000000 | 0.000% | 0.937316 | 1.000842 | no |
| `duplicate-long-quoted-after-1` | `consume` | 0.995921 | -0.408% | 0.990599 | 1.003706 | no |
| `duplicate-long-quoted-after-2` | `construct` | 0.949242 | -5.076% | 0.871611 | 1.078154 | no |
| `duplicate-long-quoted-after-2` | `consume` | 1.004200 | 0.420% | 0.999456 | 1.005421 | no |
| `duplicate-long-quoted-after-33` | `construct` | 0.999916 | -0.008% | 0.862651 | 1.097785 | no |
| `duplicate-long-quoted-after-33` | `consume` | 0.826499 | -17.350% | 0.811032 | 0.842056 | no |
| `duplicate-long-unterminated-after-1` | `construct` | 1.000000 | 0.000% | 0.997476 | 1.001774 | no |
| `duplicate-long-unterminated-after-1` | `consume` | 0.999629 | -0.037% | 0.968532 | 1.007395 | no |
| `duplicate-long-unterminated-after-2` | `construct` | 1.000000 | 0.000% | 0.914685 | 1.129848 | no |
| `duplicate-long-unterminated-after-2` | `consume` | 1.003492 | 0.349% | 0.995311 | 1.004319 | no |
| `duplicate-long-unterminated-after-33` | `construct` | 1.000843 | 0.084% | 0.937768 | 1.067454 | no |
| `duplicate-long-unterminated-after-33` | `consume` | 0.813205 | -18.680% | 0.798223 | 0.825591 | no |
| `duplicate-unquoted-after-1` | `construct` | 0.999158 | -0.084% | 0.934546 | 1.069143 | no |
| `duplicate-unquoted-after-1` | `consume` | 1.000045 | 0.005% | 0.996756 | 1.006239 | no |
| `duplicate-unquoted-after-33` | `construct` | 1.000084 | 0.008% | 0.928571 | 1.072392 | no |
| `duplicate-unquoted-after-33` | `consume` | 0.826046 | -17.395% | 0.807865 | 0.844894 | no |
| `duplicate-valid-after-1` | `construct` | 1.145025 | 14.503% | 0.902776 | 1.157673 | no |
| `duplicate-valid-after-1` | `consume` | 0.999368 | -0.063% | 0.949918 | 1.006196 | no |
| `duplicate-valid-after-2` | `construct` | 1.000843 | 0.084% | 1.000000 | 1.136477 | no |
| `duplicate-valid-after-2` | `consume` | 1.002771 | 0.277% | 0.998882 | 1.003793 | no |
| `duplicate-valid-after-32` | `construct` | 1.000000 | 0.000% | 0.995798 | 1.000843 | no |
| `duplicate-valid-after-32` | `consume` | 0.758332 | -24.167% | 0.755175 | 0.759624 | no |
| `duplicate-valid-after-33` | `construct` | 1.000000 | 0.000% | 0.999158 | 1.001686 | no |
| `duplicate-valid-after-33` | `consume` | 0.821211 | -17.879% | 0.810609 | 0.831526 | no |
| `duplicate-valid-after-4` | `construct` | 1.000000 | 0.000% | 0.936362 | 1.097808 | no |
| `duplicate-valid-after-4` | `consume` | 1.000259 | 0.026% | 0.973343 | 1.040476 | no |
| `duplicate-valid-after-5` | `construct` | 0.999157 | -0.084% | 0.963890 | 1.071669 | no |
| `duplicate-valid-after-5` | `consume` | 0.996141 | -0.386% | 0.969969 | 1.029499 | no |
| `syntax-equals-value-after-0` | `construct` | 0.998319 | -0.168% | 0.868230 | 1.084175 | no |
| `syntax-equals-value-after-0` | `consume` | 0.997225 | -0.278% | 0.972927 | 1.047826 | no |
| `syntax-equals-value-after-2` | `construct` | 1.070591 | 7.059% | 0.928181 | 1.156830 | no |
| `syntax-equals-value-after-2` | `consume` | 1.004103 | 0.410% | 1.002614 | 1.006178 | no |
| `syntax-equals-value-after-33` | `construct` | 1.000000 | 0.000% | 0.999073 | 1.070826 | no |
| `syntax-equals-value-after-33` | `consume` | 0.832138 | -16.786% | 0.818038 | 0.841417 | no |
| `syntax-equals-value-after-4` | `construct` | 0.998316 | -0.168% | 0.540046 | 1.064082 | no |
| `syntax-equals-value-after-4` | `consume` | 1.024252 | 2.425% | 1.014202 | 1.078511 | no |
| `syntax-flag-after-0` | `construct` | 1.000000 | 0.000% | 0.923451 | 1.077572 | no |
| `syntax-flag-after-0` | `consume` | 0.981945 | -1.806% | 0.959492 | 1.003773 | no |
| `syntax-flag-after-2` | `construct` | 1.002532 | 0.253% | 0.600389 | 1.134680 | no |
| `syntax-flag-after-2` | `consume` | 1.006341 | 0.634% | 0.998415 | 1.012399 | no |
| `syntax-flag-after-33` | `construct` | 1.000000 | 0.000% | 0.944526 | 1.000000 | no |
| `syntax-flag-after-33` | `consume` | 0.822076 | -17.792% | 0.813242 | 0.840139 | no |
| `syntax-flag-after-4` | `construct` | 1.000000 | 0.000% | 0.938316 | 1.080103 | no |
| `syntax-flag-after-4` | `consume` | 1.023465 | 2.346% | 1.010846 | 1.077290 | no |
| `syntax-unique-tail-after-0` | `construct` | 1.000843 | 0.084% | 0.998316 | 1.081027 | no |
| `syntax-unique-tail-after-0` | `consume` | 1.002535 | 0.254% | 0.978770 | 1.026860 | no |
| `syntax-unique-tail-after-2` | `construct` | 1.000843 | 0.084% | 0.887744 | 1.105190 | no |
| `syntax-unique-tail-after-2` | `consume` | 0.987480 | -1.252% | 0.985598 | 1.017329 | no |
| `syntax-unique-tail-after-33` | `construct` | 1.000845 | 0.084% | 0.936929 | 1.132155 | no |
| `syntax-unique-tail-after-33` | `consume` | 0.834939 | -16.506% | 0.813481 | 0.837840 | no |
| `syntax-unique-tail-after-4` | `construct` | 1.054806 | 5.481% | 0.867220 | 1.131440 | no |
| `syntax-unique-tail-after-4` | `consume` | 0.903441 | -9.656% | 0.888030 | 0.921341 | no |

Diagnostic flags: 0. A flag means the median ratio is above 1.05 and the bootstrap lower endpoint is above 1.0; it is not an adoption decision.

Process p50 spread flags above 5%: 75.

## Callgrind profile diagnostics

The parser checked 312 positive dumps and 312 empty termination dumps over the events `Ir, Bc, Bcm, Bi, Bim`.

Exact owner qualification succeeded for 312 profiles and failed for 0. Failures remain in the JSON evidence and are not treated as successful attribution.

The selected self-function census covers checked-iterator, IterState, drop, allocator-name, and lexical-matching symbols. The allocator group is lexical name matching only; it does not count allocator API calls. Inclusive graph costs overlap, so no fractions are reported.

The full receipt identities, raw Callgrind conservation records, owner qualification state, and selected symbol rows are retained in `analysis.json`.
