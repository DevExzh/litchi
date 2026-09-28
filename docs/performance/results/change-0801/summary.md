# 0801 attribute-boundary native and Callgrind analysis

This document is an offline replay of 936 native children and 312 Callgrind children. Native process p50 values use the frozen nearest-rank rule; Callgrind counters are guest diagnostics and do not measure native latency.

## Native paired diagnostics

| Case | Mode | Median after/before p50 | Change | CI low | CI high | Diagnostic regression |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| `distinct-0` | `construct` | 0.923272 | -7.673% | 0.921706 | 0.924895 | no |
| `distinct-0` | `consume` | 1.075420 | 7.542% | 1.062726 | 1.077057 | yes |
| `distinct-1` | `construct` | 0.923337 | -7.666% | 0.919328 | 1.117268 | no |
| `distinct-1` | `consume` | 0.610504 | -38.950% | 0.594645 | 0.714536 | no |
| `distinct-16` | `construct` | 0.923337 | -7.666% | 0.920091 | 1.042929 | no |
| `distinct-16` | `consume` | 1.680343 | 68.034% | 1.667178 | 1.724234 | yes |
| `distinct-17` | `construct` | 0.924115 | -7.589% | 0.922559 | 1.041182 | no |
| `distinct-17` | `consume` | 1.634389 | 63.439% | 1.619928 | 1.661384 | yes |
| `distinct-2` | `construct` | 0.925738 | -7.426% | 0.924179 | 1.164847 | no |
| `distinct-2` | `consume` | 0.806294 | -19.371% | 0.790086 | 0.806685 | no |
| `distinct-3` | `construct` | 0.923402 | -7.660% | 0.920946 | 1.000778 | no |
| `distinct-3` | `consume` | 0.908208 | -9.179% | 0.894137 | 0.922479 | no |
| `distinct-32` | `construct` | 0.923337 | -7.666% | 0.921717 | 1.004988 | no |
| `distinct-32` | `consume` | 1.599407 | 59.941% | 1.582182 | 1.667345 | yes |
| `distinct-33` | `construct` | 0.924179 | -7.582% | 0.921782 | 1.042867 | no |
| `distinct-33` | `consume` | 0.845074 | -15.493% | 0.836047 | 0.850544 | no |
| `distinct-4` | `construct` | 0.922559 | -7.744% | 0.920942 | 1.041530 | no |
| `distinct-4` | `consume` | 1.611177 | 61.118% | 1.594146 | 1.638939 | yes |
| `distinct-5` | `construct` | 0.923337 | -7.666% | 0.922481 | 1.038007 | no |
| `distinct-5` | `consume` | 1.327504 | 32.750% | 1.296225 | 1.410613 | yes |
| `distinct-64` | `construct` | 0.922559 | -7.744% | 0.921717 | 0.923402 | no |
| `distinct-64` | `consume` | 1.001356 | 0.136% | 0.994029 | 1.010791 | no |
| `distinct-8` | `construct` | 0.924179 | -7.582% | 0.923259 | 1.123810 | no |
| `distinct-8` | `consume` | 1.466947 | 46.695% | 1.442879 | 1.497196 | yes |
| `distinct-9` | `construct` | 0.924245 | -7.575% | 0.922559 | 1.000779 | no |
| `distinct-9` | `consume` | 1.401737 | 40.174% | 1.372058 | 1.443409 | yes |
| `duplicate-long-quoted-after-1` | `construct` | 0.923466 | -7.653% | 0.920943 | 1.043710 | no |
| `duplicate-long-quoted-after-1` | `consume` | 0.669080 | -33.092% | 0.646842 | 0.674357 | no |
| `duplicate-long-quoted-after-2` | `construct` | 0.924115 | -7.589% | 0.922494 | 0.999158 | no |
| `duplicate-long-quoted-after-2` | `consume` | 0.820570 | -17.943% | 0.814813 | 0.828455 | no |
| `duplicate-long-quoted-after-33` | `construct` | 0.922559 | -7.744% | 0.888450 | 0.925801 | no |
| `duplicate-long-quoted-after-33` | `consume` | 0.552602 | -44.740% | 0.484738 | 0.558279 | no |
| `duplicate-long-unterminated-after-1` | `construct` | 0.924179 | -7.582% | 0.921784 | 1.165964 | no |
| `duplicate-long-unterminated-after-1` | `consume` | 0.658377 | -34.162% | 0.641220 | 0.668998 | no |
| `duplicate-long-unterminated-after-2` | `construct` | 1.004968 | 0.497% | 0.923466 | 1.161616 | no |
| `duplicate-long-unterminated-after-2` | `consume` | 0.829253 | -17.075% | 0.815069 | 0.841474 | no |
| `duplicate-long-unterminated-after-33` | `construct` | 0.922559 | -7.744% | 0.921784 | 1.151258 | no |
| `duplicate-long-unterminated-after-33` | `consume` | 0.554517 | -44.548% | 0.548345 | 0.560546 | no |
| `duplicate-unquoted-after-1` | `construct` | 0.922624 | -7.738% | 0.921784 | 1.049747 | no |
| `duplicate-unquoted-after-1` | `consume` | 0.657632 | -34.237% | 0.647174 | 0.668548 | no |
| `duplicate-unquoted-after-33` | `construct` | 0.924179 | -7.582% | 0.921008 | 1.165825 | no |
| `duplicate-unquoted-after-33` | `consume` | 0.853292 | -14.671% | 0.840481 | 0.873300 | no |
| `duplicate-valid-after-1` | `construct` | 0.923401 | -7.660% | 0.920942 | 1.125499 | no |
| `duplicate-valid-after-1` | `consume` | 0.656690 | -34.331% | 0.650714 | 0.665084 | no |
| `duplicate-valid-after-2` | `construct` | 0.922559 | -7.744% | 0.889410 | 1.041183 | no |
| `duplicate-valid-after-2` | `consume` | 0.823708 | -17.629% | 0.810587 | 0.837124 | no |
| `duplicate-valid-after-32` | `construct` | 0.922692 | -7.731% | 0.921008 | 1.036071 | no |
| `duplicate-valid-after-32` | `consume` | 0.843526 | -15.647% | 0.829657 | 0.861269 | no |
| `duplicate-valid-after-33` | `construct` | 0.924958 | -7.504% | 0.920875 | 1.156693 | no |
| `duplicate-valid-after-33` | `consume` | 0.846747 | -15.325% | 0.840066 | 0.852118 | no |
| `duplicate-valid-after-4` | `construct` | 0.923402 | -7.660% | 0.921784 | 0.924306 | no |
| `duplicate-valid-after-4` | `consume` | 1.447408 | 44.741% | 1.421704 | 1.487780 | yes |
| `duplicate-valid-after-5` | `construct` | 0.924242 | -7.576% | 0.922495 | 1.168076 | no |
| `duplicate-valid-after-5` | `consume` | 1.173103 | 17.310% | 1.144606 | 1.198224 | yes |
| `syntax-equals-value-after-0` | `construct` | 0.923337 | -7.666% | 0.921784 | 0.925022 | no |
| `syntax-equals-value-after-0` | `consume` | 1.040774 | 4.077% | 1.017095 | 1.078795 | no |
| `syntax-equals-value-after-2` | `construct` | 0.925801 | -7.420% | 0.923272 | 1.002525 | no |
| `syntax-equals-value-after-2` | `consume` | 0.972962 | -2.704% | 0.964832 | 0.980431 | no |
| `syntax-equals-value-after-33` | `construct` | 1.040342 | 4.034% | 0.921717 | 1.175379 | no |
| `syntax-equals-value-after-33` | `consume` | 0.843290 | -15.671% | 0.799966 | 0.859422 | no |
| `syntax-equals-value-after-4` | `construct` | 0.925022 | -7.498% | 0.923337 | 1.158959 | no |
| `syntax-equals-value-after-4` | `consume` | 1.636439 | 63.644% | 1.614249 | 1.668411 | yes |
| `syntax-flag-after-0` | `construct` | 0.923421 | -7.658% | 0.922559 | 1.044613 | no |
| `syntax-flag-after-0` | `consume` | 1.045865 | 4.586% | 1.010089 | 1.065937 | no |
| `syntax-flag-after-2` | `construct` | 1.001582 | 0.158% | 0.924180 | 1.135395 | no |
| `syntax-flag-after-2` | `consume` | 0.961742 | -3.826% | 0.952555 | 0.981642 | no |
| `syntax-flag-after-33` | `construct` | 0.922559 | -7.744% | 0.921497 | 0.923466 | no |
| `syntax-flag-after-33` | `consume` | 0.847931 | -15.207% | 0.837362 | 0.854237 | no |
| `syntax-flag-after-4` | `construct` | 0.923337 | -7.666% | 0.921717 | 1.036216 | no |
| `syntax-flag-after-4` | `consume` | 1.626210 | 62.621% | 1.619037 | 1.635019 | yes |
| `syntax-unique-tail-after-0` | `construct` | 0.925022 | -7.498% | 0.923337 | 1.122193 | no |
| `syntax-unique-tail-after-0` | `consume` | 0.820798 | -17.920% | 0.777056 | 0.838339 | no |
| `syntax-unique-tail-after-2` | `construct` | 0.922624 | -7.738% | 0.921784 | 1.110455 | no |
| `syntax-unique-tail-after-2` | `consume` | 0.953756 | -4.624% | 0.933742 | 0.978573 | no |
| `syntax-unique-tail-after-33` | `construct` | 0.922559 | -7.744% | 0.921008 | 1.041049 | no |
| `syntax-unique-tail-after-33` | `consume` | 0.871407 | -12.859% | 0.861912 | 0.877856 | no |
| `syntax-unique-tail-after-4` | `construct` | 0.998076 | -0.192% | 0.921643 | 1.159121 | no |
| `syntax-unique-tail-after-4` | `consume` | 1.367754 | 36.775% | 1.348996 | 1.382901 | yes |

Diagnostic flags: 13. A flag means the median ratio is above 1.05 and the bootstrap lower endpoint is above 1.0; it is not an adoption decision.

Process p50 spread flags above 5%: 50.

## Callgrind profile diagnostics

The parser checked 312 positive dumps and 312 empty termination dumps over the events `Ir, Bc, Bcm, Bi, Bim`.

Exact owner qualification succeeded for 312 profiles and failed for 0. Failures remain in the JSON evidence and are not treated as successful attribution.

The selected self-function census covers checked-iterator, IterState, drop, allocator-name, and lexical-matching symbols. The allocator group is lexical name matching only; it does not count allocator API calls. Inclusive graph costs overlap, so no fractions are reported.

The full receipt identities, raw Callgrind conservation records, owner qualification state, and selected symbol rows are retained in `analysis.json`.
