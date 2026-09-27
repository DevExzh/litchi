# 0788 heaptrack allocation attribution

This is a retained heaptrack diagnostic for the rejected cached-Part candidate.
Heaptrack intercepted-allocation peaks are reported separately from RSS and native operation timing.
Merged backtraces were disabled because heaptrack warns that merged peak consumption is inaccurate.

## Summary

- Reports: 16; samples per child: 30; repeats: 2.
- Allocation fields absent from a heaptrack version remain unavailable; the parser does not substitute zero.
- Human-readable heaptrack byte fields retain their raw rounded spelling below; normalized numbers remain in JSON for comparison.
- The report supplies attribution evidence only and makes no causal claim.

## Children

| shape | state | floor | width | leg | repeat | allocations | total allocated | peak live | leaked | temporary | flame stacks |
|---|---|---:|---:|---|---:|---:|---:|---:|---:|---:|---:|
| small | fresh | 0 | 4 | after | 0 | 94486 (549337/s) | 108478093B (histogram) | ~981.64K | ~2.00K | 38329 (222843/s) | 428 |
| small | fresh | 0 | 4 | after | 1 | 94482 (543000/s) | 108477709B (histogram) | ~982.06K | ~2.00K | 38322 (220241/s) | 428 |
| small | fresh | 0 | 4 | before | 0 | 94484 (546150/s) | 108477903B (histogram) | ~981.98K | ~2.00K | 38320 (221502/s) | 428 |
| small | fresh | 0 | 4 | before | 1 | 94484 (533807/s) | 108477903B (histogram) | ~980.16K | ~2.00K | 38327 (216536/s) | 428 |
| small | primed | 0 | 4 | after | 0 | 132666 (728934/s) | 108937894B (histogram) | ~1.01M | ~2.00K | 68806 (378054/s) | 450 |
| small | primed | 0 | 4 | after | 1 | 132667 (724956/s) | 108937990B (histogram) | ~982.35K | ~2.00K | 68815 (376038/s) | 450 |
| small | primed | 0 | 4 | before | 0 | 134931 (674655/s) | 109228464B (histogram) | ~981.54K | ~2.00K | 68910 (344550/s) | 478 |
| small | primed | 0 | 4 | before | 1 | 134935 (688443/s) | 109228848B (histogram) | ~980.04K | ~2.00K | 68913 (351596/s) | 478 |
| small | primed | 0 | 32 | after | 0 | 139548 (596358/s) | 109455088B (histogram) | ~1.13M | ~7.28K | 69435 (296730/s) | 433 |
| small | primed | 0 | 32 | after | 1 | 139533 (625708/s) | 109447360B (histogram) | ~1.06M | ~7.28K | 69226 (310430/s) | 433 |
| small | primed | 0 | 32 | before | 0 | 148536 (538173/s) | 110183178B (histogram) | ~1.09M | ~7.28K | 70013 (253670/s) | 447 |
| small | primed | 0 | 32 | before | 1 | 148498 (536093/s) | 110160394B (histogram) | ~1.09M | ~7.28K | 70000 (252707/s) | 447 |
| small | primed | 65536 | 4 | after | 0 | 130355 (917992/s) | 108640270B (histogram) | ~805.85K | 544B | 68717 (483922/s) | 423 |
| small | primed | 65536 | 4 | after | 1 | 130355 (924503/s) | 108640270B (histogram) | ~805.85K | 544B | 68717 (487354/s) | 423 |
| small | primed | 65536 | 4 | before | 0 | 130355 (924503/s) | 108640272B (histogram) | ~805.85K | 544B | 68717 (487354/s) | 423 |
| small | primed | 65536 | 4 | before | 1 | 130355 (937805/s) | 108640272B (histogram) | ~805.85K | 544B | 68717 (494366/s) | 423 |

## ELF allocated section deltas

Values come from retained `readelf -W -S` output; only allocated sections are classified.

| binary | text | rodata | data | bss | eh_frame | total |
|---|---:|---:|---:|---:|---:|---:|
| native | 2864 | 88 | 0 | 1056 | 96 | 4104 |
| memory | 2864 | 88 | 0 | 1056 | 96 | 4104 |

## Exact peak-cost stack pairs

The totals below sum the unmerged peak-cost flamegraph lines and cross-check against the rounded print summary. They describe intercepted allocation peak attribution.

| shape | state | floor | width | repeat | before stack cost | after stack cost | delta |
|---|---|---:|---:|---:|---:|---:|---:|
| small | fresh | 0 | 4 | 0 | 981976 | 981644 | -332 |
| small | fresh | 0 | 4 | 1 | 980161 | 982063 | +1902 |
| small | primed | 0 | 4 | 0 | 981543 | 1008515 | +26972 |
| small | primed | 0 | 4 | 1 | 980042 | 982353 | +2311 |
| small | primed | 0 | 32 | 0 | 1091806 | 1125676 | +33870 |
| small | primed | 0 | 32 | 1 | 1087285 | 1057567 | -29718 |
| small | primed | 65536 | 4 | 0 | 805854 | 805853 | -1 |
| small | primed | 65536 | 4 | 1 | 805854 | 805853 | -1 |

## Top peak stack evidence

The first relevant frame is retained as a bounded attribution label; it does not establish causation.

| shape | state | floor | width | leg | repeat | top cost | first relevant frame |
|---|---|---:|---:|---|---:|---:|---|
| small | fresh | 0 | 4 | after | 0 | 147456 | litchi_perf_execution::build_corpus::h63151714ab2f248a (main.rs) |
| small | fresh | 0 | 4 | after | 1 | 147456 | litchi_perf_execution::build_corpus::h63151714ab2f248a (main.rs) |
| small | fresh | 0 | 4 | before | 0 | 147456 | litchi_perf_execution::build_corpus::h63151714ab2f248a (main.rs) |
| small | fresh | 0 | 4 | before | 1 | 147456 | litchi_perf_execution::build_corpus::h63151714ab2f248a (main.rs) |
| small | primed | 0 | 4 | after | 0 | 190208 | start_thread (pthread_create.c) |
| small | primed | 0 | 4 | after | 1 | 147456 | litchi_perf_execution::build_corpus::h63151714ab2f248a (main.rs) |
| small | primed | 0 | 4 | before | 0 | 147456 | litchi_perf_execution::build_corpus::h63151714ab2f248a (main.rs) |
| small | primed | 0 | 4 | before | 1 | 147456 | litchi_perf_execution::build_corpus::h63151714ab2f248a (main.rs) |
| small | primed | 0 | 32 | after | 0 | 237760 | start_thread (pthread_create.c) |
| small | primed | 0 | 32 | after | 1 | 190208 | start_thread (pthread_create.c) |
| small | primed | 0 | 32 | before | 0 | 237760 | start_thread (pthread_create.c) |
| small | primed | 0 | 32 | before | 1 | 190208 | start_thread (pthread_create.c) |
| small | primed | 65536 | 4 | after | 0 | 147456 | litchi_perf_execution::build_corpus::h63151714ab2f248a (main.rs) |
| small | primed | 65536 | 4 | after | 1 | 147456 | litchi_perf_execution::build_corpus::h63151714ab2f248a (main.rs) |
| small | primed | 65536 | 4 | before | 0 | 147456 | litchi_perf_execution::build_corpus::h63151714ab2f248a (main.rs) |
| small | primed | 65536 | 4 | before | 1 | 147456 | litchi_perf_execution::build_corpus::h63151714ab2f248a (main.rs) |

The native W4 primed floor-0 pair deltas are +26972 B, +2311 B across repeats. These are allocation-profile deltas and establish no RSS cause.
The intercepted allocation peak is not a native RSS high-water mark and is not pooled with native timing or RSS.
