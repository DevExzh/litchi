# 0802 attribute-boundary native and Callgrind analysis

This document is an offline replay of 936 native children and 312 Callgrind children. Native process p50 values use the frozen nearest-rank rule; Callgrind counters are guest diagnostics and do not measure native latency.

## Native paired diagnostics

| Case | Mode | Median after/before p50 | Change | CI low | CI high | Diagnostic regression |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| `distinct-0` | `construct` | 3.177667 | 217.767% | 2.994513 | 3.372681 | yes |
| `distinct-0` | `consume` | 0.675011 | -32.499% | 0.674321 | 0.676301 | no |
| `distinct-1` | `construct` | 3.369924 | 236.992% | 3.217355 | 3.372681 | yes |
| `distinct-1` | `consume` | 0.799373 | -20.063% | 0.795323 | 0.804328 | no |
| `distinct-16` | `construct` | 3.288560 | 228.856% | 3.061138 | 3.406879 | yes |
| `distinct-16` | `consume` | 1.275080 | 27.508% | 1.251021 | 1.287138 | yes |
| `distinct-17` | `construct` | 3.291125 | 229.112% | 2.985285 | 3.370684 | yes |
| `distinct-17` | `consume` | 1.197851 | 19.785% | 1.189062 | 1.212027 | yes |
| `distinct-2` | `construct` | 3.223630 | 222.363% | 3.177154 | 3.413682 | yes |
| `distinct-2` | `consume` | 0.928650 | -7.135% | 0.908952 | 0.939217 | no |
| `distinct-3` | `construct` | 3.176686 | 217.669% | 2.969706 | 3.374934 | yes |
| `distinct-3` | `consume` | 1.024019 | 2.402% | 1.014697 | 1.027424 | no |
| `distinct-32` | `construct` | 3.252717 | 225.272% | 2.682357 | 3.414983 | yes |
| `distinct-32` | `consume` | 1.379962 | 37.996% | 1.375706 | 1.394342 | yes |
| `distinct-33` | `construct` | 3.204559 | 220.456% | 3.135477 | 3.331472 | yes |
| `distinct-33` | `consume` | 0.769327 | -23.067% | 0.725713 | 0.773088 | no |
| `distinct-4` | `construct` | 3.369217 | 236.922% | 3.291781 | 3.379104 | yes |
| `distinct-4` | `consume` | 1.302379 | 30.238% | 1.287646 | 1.323957 | yes |
| `distinct-5` | `construct` | 3.209301 | 220.930% | 3.133809 | 3.404882 | yes |
| `distinct-5` | `consume` | 1.027207 | 2.721% | 1.013039 | 1.039386 | no |
| `distinct-64` | `construct` | 3.220067 | 222.007% | 2.935387 | 3.373212 | yes |
| `distinct-64` | `consume` | 1.107857 | 10.786% | 1.095758 | 1.130985 | yes |
| `distinct-8` | `construct` | 3.296571 | 229.657% | 3.177794 | 3.372368 | yes |
| `distinct-8` | `consume` | 1.061591 | 6.159% | 1.056122 | 1.094124 | yes |
| `distinct-9` | `construct` | 3.371212 | 237.121% | 3.285714 | 3.374452 | yes |
| `distinct-9` | `consume` | 1.001389 | 0.139% | 0.991255 | 1.013284 | no |
| `duplicate-long-quoted-after-1` | `construct` | 3.371526 | 237.153% | 3.054402 | 3.377747 | yes |
| `duplicate-long-quoted-after-1` | `consume` | 0.761903 | -23.810% | 0.753415 | 0.763485 | no |
| `duplicate-long-quoted-after-2` | `construct` | 3.287008 | 228.701% | 2.975521 | 3.408333 | yes |
| `duplicate-long-quoted-after-2` | `consume` | 0.908842 | -9.116% | 0.893748 | 0.922770 | no |
| `duplicate-long-quoted-after-33` | `construct` | 3.220539 | 222.054% | 3.209088 | 3.301538 | yes |
| `duplicate-long-quoted-after-33` | `consume` | 0.687584 | -31.242% | 0.684380 | 0.693997 | no |
| `duplicate-long-unterminated-after-1` | `construct` | 3.373524 | 237.352% | 3.255276 | 3.412110 | yes |
| `duplicate-long-unterminated-after-1` | `consume` | 0.771182 | -22.882% | 0.758099 | 0.781930 | no |
| `duplicate-long-unterminated-after-2` | `construct` | 3.370770 | 237.077% | 3.132136 | 3.375737 | yes |
| `duplicate-long-unterminated-after-2` | `consume` | 0.912318 | -8.768% | 0.898211 | 0.914810 | no |
| `duplicate-long-unterminated-after-33` | `construct` | 3.299015 | 229.901% | 3.105039 | 3.414458 | yes |
| `duplicate-long-unterminated-after-33` | `consume` | 0.692650 | -30.735% | 0.679441 | 0.703557 | no |
| `duplicate-unquoted-after-1` | `construct` | 3.295329 | 229.533% | 2.995301 | 3.413153 | yes |
| `duplicate-unquoted-after-1` | `consume` | 0.759602 | -24.040% | 0.753843 | 0.768451 | no |
| `duplicate-unquoted-after-33` | `construct` | 3.373211 | 237.321% | 3.215325 | 3.417369 | yes |
| `duplicate-unquoted-after-33` | `consume` | 1.059149 | 5.915% | 1.035592 | 1.080875 | yes |
| `duplicate-valid-after-1` | `construct` | 3.368687 | 236.869% | 3.179626 | 3.375740 | yes |
| `duplicate-valid-after-1` | `consume` | 0.762491 | -23.751% | 0.758171 | 0.798917 | no |
| `duplicate-valid-after-2` | `construct` | 3.366694 | 236.669% | 3.133020 | 3.377829 | yes |
| `duplicate-valid-after-2` | `consume` | 0.921092 | -7.891% | 0.908887 | 0.934139 | no |
| `duplicate-valid-after-32` | `construct` | 3.220911 | 222.091% | 3.104667 | 3.409933 | yes |
| `duplicate-valid-after-32` | `consume` | 0.735630 | -26.437% | 0.726767 | 0.740289 | no |
| `duplicate-valid-after-33` | `construct` | 3.208783 | 220.878% | 3.137147 | 3.374083 | yes |
| `duplicate-valid-after-33` | `consume` | 1.060435 | 6.043% | 1.054929 | 1.062753 | yes |
| `duplicate-valid-after-4` | `construct` | 3.373212 | 237.321% | 3.183425 | 3.412776 | yes |
| `duplicate-valid-after-4` | `consume` | 1.178978 | 17.898% | 1.168146 | 1.214728 | yes |
| `duplicate-valid-after-5` | `construct` | 3.406122 | 240.612% | 3.298254 | 3.451560 | yes |
| `duplicate-valid-after-5` | `consume` | 0.971371 | -2.863% | 0.966721 | 0.996125 | no |
| `syntax-equals-value-after-0` | `construct` | 3.222597 | 222.260% | 3.098592 | 3.377744 | yes |
| `syntax-equals-value-after-0` | `consume` | 1.347073 | 34.707% | 1.296092 | 1.359071 | yes |
| `syntax-equals-value-after-2` | `construct` | 3.293954 | 229.395% | 3.171467 | 3.375295 | yes |
| `syntax-equals-value-after-2` | `consume` | 1.072755 | 7.275% | 1.052862 | 1.082765 | yes |
| `syntax-equals-value-after-33` | `construct` | 3.368687 | 236.869% | 3.170063 | 3.449032 | yes |
| `syntax-equals-value-after-33` | `consume` | 1.050813 | 5.081% | 1.047279 | 1.057305 | yes |
| `syntax-equals-value-after-4` | `construct` | 3.220725 | 222.072% | 3.178393 | 3.298169 | yes |
| `syntax-equals-value-after-4` | `consume` | 1.356291 | 35.629% | 1.349457 | 1.362999 | yes |
| `syntax-flag-after-0` | `construct` | 3.295953 | 229.595% | 3.174961 | 3.405838 | yes |
| `syntax-flag-after-0` | `consume` | 1.276295 | 27.629% | 1.243891 | 1.315446 | yes |
| `syntax-flag-after-2` | `construct` | 3.301096 | 230.110% | 2.934145 | 3.449035 | yes |
| `syntax-flag-after-2` | `consume` | 1.076772 | 7.677% | 1.059236 | 1.080236 | yes |
| `syntax-flag-after-33` | `construct` | 3.299884 | 229.988% | 3.220911 | 3.413153 | yes |
| `syntax-flag-after-33` | `consume` | 1.062360 | 6.236% | 1.055413 | 1.073918 | yes |
| `syntax-flag-after-4` | `construct` | 3.294697 | 229.470% | 3.101888 | 3.372926 | yes |
| `syntax-flag-after-4` | `consume` | 1.350458 | 35.046% | 1.300301 | 1.360854 | yes |
| `syntax-unique-tail-after-0` | `construct` | 3.292929 | 229.293% | 3.137931 | 3.410333 | yes |
| `syntax-unique-tail-after-0` | `consume` | 1.003958 | 0.396% | 1.000437 | 1.011211 | no |
| `syntax-unique-tail-after-2` | `construct` | 3.374055 | 237.406% | 3.288575 | 3.456540 | yes |
| `syntax-unique-tail-after-2` | `consume` | 1.045561 | 4.556% | 1.038566 | 1.061434 | no |
| `syntax-unique-tail-after-33` | `construct` | 3.180997 | 218.100% | 3.129898 | 3.406566 | yes |
| `syntax-unique-tail-after-33` | `consume` | 1.064310 | 6.431% | 1.061422 | 1.074081 | yes |
| `syntax-unique-tail-after-4` | `construct` | 3.374980 | 237.498% | 3.108229 | 3.445664 | yes |
| `syntax-unique-tail-after-4` | `consume` | 1.176228 | 17.623% | 1.166360 | 1.186275 | yes |

Diagnostic flags: 58. A flag means the median ratio is above 1.05 and the bootstrap lower endpoint is above 1.0; it is not an adoption decision.

Process p50 spread flags above 5%: 66.

## Callgrind profile diagnostics

The parser checked 312 positive dumps and 312 empty termination dumps over the events `Ir, Bc, Bcm, Bi, Bim`.

Exact owner qualification succeeded for 312 profiles and failed for 0. Failures remain in the JSON evidence and are not treated as successful attribution.

The selected self-function census covers checked-iterator, IterState, drop, allocator-name, and lexical-matching symbols. The allocator group is lexical name matching only; it does not count allocator API calls. Inclusive graph costs overlap, so no fractions are reported.

The full receipt identities, raw Callgrind conservation records, owner qualification state, and selected symbol rows are retained in `analysis.json`.
