# 0803 attribute-boundary native and Callgrind analysis

This document is an offline replay of 936 native children and 312 Callgrind children. Native process p50 values use the frozen nearest-rank rule; Callgrind counters are guest diagnostics and do not measure native latency.

## Native paired diagnostics

| Case | Mode | Median after/before p50 | Change | CI low | CI high | Diagnostic regression |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| `distinct-0` | `construct` | 0.314006 | -68.599% | 0.296007 | 0.338071 | no |
| `distinct-0` | `consume` | 0.519075 | -48.093% | 0.518323 | 0.527919 | no |
| `distinct-1` | `construct` | 0.317911 | -68.209% | 0.297000 | 0.380322 | no |
| `distinct-1` | `consume` | 0.761308 | -23.869% | 0.761008 | 0.768881 | no |
| `distinct-16` | `construct` | 0.314945 | -68.505% | 0.296513 | 0.356741 | no |
| `distinct-16` | `consume` | 0.968398 | -3.160% | 0.925878 | 0.987255 | no |
| `distinct-17` | `construct` | 0.337252 | -66.275% | 0.307687 | 0.382913 | no |
| `distinct-17` | `consume` | 0.985093 | -1.491% | 0.970488 | 1.002982 | no |
| `distinct-2` | `construct` | 0.326026 | -67.397% | 0.295570 | 0.358550 | no |
| `distinct-2` | `consume` | 0.868048 | -13.195% | 0.863809 | 0.873600 | no |
| `distinct-3` | `construct` | 0.314927 | -68.507% | 0.303380 | 0.361615 | no |
| `distinct-3` | `consume` | 0.926370 | -7.363% | 0.921349 | 0.933922 | no |
| `distinct-32` | `construct` | 0.303493 | -69.651% | 0.295144 | 0.335723 | no |
| `distinct-32` | `consume` | 0.999337 | -0.066% | 0.984156 | 1.009495 | no |
| `distinct-33` | `construct` | 0.307942 | -69.206% | 0.294952 | 0.952868 | no |
| `distinct-33` | `consume` | 0.997026 | -0.297% | 0.989629 | 1.002884 | no |
| `distinct-4` | `construct` | 0.340265 | -65.973% | 0.295828 | 0.367670 | no |
| `distinct-4` | `consume` | 0.933885 | -6.612% | 0.910526 | 0.954013 | no |
| `distinct-5` | `construct` | 0.315243 | -68.476% | 0.302921 | 0.334730 | no |
| `distinct-5` | `consume` | 0.964053 | -3.595% | 0.919000 | 0.988150 | no |
| `distinct-64` | `construct` | 0.296817 | -70.318% | 0.296084 | 0.404045 | no |
| `distinct-64` | `consume` | 0.988635 | -1.137% | 0.973647 | 1.010737 | no |
| `distinct-8` | `construct` | 0.323974 | -67.603% | 0.295395 | 0.386280 | no |
| `distinct-8` | `consume` | 0.980963 | -1.904% | 0.941189 | 1.009662 | no |
| `distinct-9` | `construct` | 0.303661 | -69.634% | 0.295717 | 0.374536 | no |
| `distinct-9` | `consume` | 0.962445 | -3.755% | 0.955368 | 0.982714 | no |
| `duplicate-long-quoted-after-1` | `construct` | 0.296307 | -70.369% | 0.295099 | 0.377580 | no |
| `duplicate-long-quoted-after-1` | `consume` | 0.837057 | -16.294% | 0.834029 | 0.845487 | no |
| `duplicate-long-quoted-after-2` | `construct` | 0.339477 | -66.052% | 0.307243 | 0.374880 | no |
| `duplicate-long-quoted-after-2` | `consume` | 0.897531 | -10.247% | 0.883793 | 0.899226 | no |
| `duplicate-long-quoted-after-33` | `construct` | 0.297000 | -70.300% | 0.295568 | 0.351885 | no |
| `duplicate-long-quoted-after-33` | `consume` | 1.001379 | 0.138% | 0.987066 | 1.046967 | no |
| `duplicate-long-unterminated-after-1` | `construct` | 0.318762 | -68.124% | 0.222096 | 0.386966 | no |
| `duplicate-long-unterminated-after-1` | `consume` | 0.831341 | -16.866% | 0.818662 | 0.837820 | no |
| `duplicate-long-unterminated-after-2` | `construct` | 0.335838 | -66.416% | 0.297160 | 0.403200 | no |
| `duplicate-long-unterminated-after-2` | `consume` | 0.893810 | -10.619% | 0.821256 | 0.901932 | no |
| `duplicate-long-unterminated-after-33` | `construct` | 0.296594 | -70.341% | 0.295417 | 0.330676 | no |
| `duplicate-long-unterminated-after-33` | `consume` | 0.997431 | -0.257% | 0.980127 | 1.015727 | no |
| `duplicate-unquoted-after-1` | `construct` | 0.322519 | -67.748% | 0.296676 | 0.405092 | no |
| `duplicate-unquoted-after-1` | `consume` | 0.835601 | -16.440% | 0.830072 | 0.837662 | no |
| `duplicate-unquoted-after-33` | `construct` | 0.296528 | -70.347% | 0.295284 | 0.341763 | no |
| `duplicate-unquoted-after-33` | `consume` | 1.005577 | 0.558% | 0.991574 | 1.018639 | no |
| `duplicate-valid-after-1` | `construct` | 0.318061 | -68.194% | 0.296676 | 0.368026 | no |
| `duplicate-valid-after-1` | `consume` | 0.835355 | -16.465% | 0.831962 | 0.847003 | no |
| `duplicate-valid-after-2` | `construct` | 0.372513 | -62.749% | 0.296493 | 0.427227 | no |
| `duplicate-valid-after-2` | `consume` | 0.897073 | -10.293% | 0.892362 | 0.901812 | no |
| `duplicate-valid-after-32` | `construct` | 0.313782 | -68.622% | 0.296188 | 0.335145 | no |
| `duplicate-valid-after-32` | `consume` | 0.990417 | -0.958% | 0.951516 | 0.999320 | no |
| `duplicate-valid-after-33` | `construct` | 0.318683 | -68.132% | 0.310490 | 0.387344 | no |
| `duplicate-valid-after-33` | `consume` | 0.992502 | -0.750% | 0.977471 | 1.023733 | no |
| `duplicate-valid-after-4` | `construct` | 0.310309 | -68.969% | 0.297548 | 0.357729 | no |
| `duplicate-valid-after-4` | `consume` | 0.973916 | -2.608% | 0.953567 | 0.983606 | no |
| `duplicate-valid-after-5` | `construct` | 0.348969 | -65.103% | 0.303439 | 0.386685 | no |
| `duplicate-valid-after-5` | `consume` | 0.977154 | -2.285% | 0.960101 | 1.007242 | no |
| `syntax-equals-value-after-0` | `construct` | 0.322231 | -67.777% | 0.303242 | 0.395073 | no |
| `syntax-equals-value-after-0` | `consume` | 0.773999 | -22.600% | 0.728791 | 0.806372 | no |
| `syntax-equals-value-after-2` | `construct` | 0.306810 | -69.319% | 0.296011 | 0.383004 | no |
| `syntax-equals-value-after-2` | `consume` | 0.907797 | -9.220% | 0.894010 | 0.911018 | no |
| `syntax-equals-value-after-33` | `construct` | 0.307460 | -69.254% | 0.294952 | 0.334822 | no |
| `syntax-equals-value-after-33` | `consume` | 0.998417 | -0.158% | 0.994070 | 1.012253 | no |
| `syntax-equals-value-after-4` | `construct` | 0.326764 | -67.324% | 0.310507 | 0.368364 | no |
| `syntax-equals-value-after-4` | `consume` | 0.946065 | -5.394% | 0.940959 | 0.977109 | no |
| `syntax-flag-after-0` | `construct` | 0.307677 | -69.232% | 0.296380 | 0.335113 | no |
| `syntax-flag-after-0` | `consume` | 0.814159 | -18.584% | 0.780290 | 0.816137 | no |
| `syntax-flag-after-2` | `construct` | 0.314909 | -68.509% | 0.296520 | 0.334381 | no |
| `syntax-flag-after-2` | `consume` | 0.907253 | -9.275% | 0.904402 | 0.909049 | no |
| `syntax-flag-after-33` | `construct` | 0.314415 | -68.558% | 0.295910 | 0.412267 | no |
| `syntax-flag-after-33` | `consume` | 0.990448 | -0.955% | 0.973713 | 1.008991 | no |
| `syntax-flag-after-4` | `construct` | 0.308256 | -69.174% | 0.295172 | 0.369696 | no |
| `syntax-flag-after-4` | `consume` | 0.943592 | -5.641% | 0.933540 | 0.977259 | no |
| `syntax-unique-tail-after-0` | `construct` | 0.340291 | -65.971% | 0.324434 | 0.389209 | no |
| `syntax-unique-tail-after-0` | `consume` | 0.818945 | -18.106% | 0.796924 | 0.840592 | no |
| `syntax-unique-tail-after-2` | `construct` | 0.315292 | -68.471% | 0.296021 | 0.369229 | no |
| `syntax-unique-tail-after-2` | `consume` | 0.915288 | -8.471% | 0.902092 | 0.921684 | no |
| `syntax-unique-tail-after-33` | `construct` | 0.318457 | -68.154% | 0.303476 | 0.373474 | no |
| `syntax-unique-tail-after-33` | `consume` | 1.002192 | 0.219% | 0.983152 | 1.007994 | no |
| `syntax-unique-tail-after-4` | `construct` | 0.311158 | -68.884% | 0.295495 | 0.385872 | no |
| `syntax-unique-tail-after-4` | `consume` | 0.965104 | -3.490% | 0.957573 | 0.997769 | no |

Diagnostic flags: 0. A flag means the median ratio is above 1.05 and the bootstrap lower endpoint is above 1.0; it is not an adoption decision.

Process p50 spread flags above 5%: 73.

## Callgrind profile diagnostics

The parser checked 312 positive dumps and 312 empty termination dumps over the events `Ir, Bc, Bcm, Bi, Bim`.

Exact owner qualification succeeded for 312 profiles and failed for 0. Failures remain in the JSON evidence and are not treated as successful attribution.

The selected self-function census covers checked-iterator, IterState, drop, allocator-name, and lexical-matching symbols. The allocator group is lexical name matching only; it does not count allocator API calls. Inclusive graph costs overlap, so no fractions are reported.

The full receipt identities, raw Callgrind conservation records, owner qualification state, and selected symbol rows are retained in `analysis.json`.
