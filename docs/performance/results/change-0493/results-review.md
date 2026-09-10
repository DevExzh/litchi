# change-0493 managed read-ahead results review

This is a bounded read-only review of the retained 0493 evidence. The pilot
has eight passing children and 24 measured samples; the formal run has 16
passing children and 480 measured samples. Canonical `verify --pilot` and
`verify` both recompute their retained analyses successfully. Every retained
report validates the managed schema, exact text oracle, source version fence,
cache diagnostics, budget release, physical trace, and allocator contract.

The evidence is bound to protocol
`d611d0b637952e21b5ba7ea5f3b1ae0575dfd96b2aeb7a76a5c0ec70ec183cb2`, source
revision `e8ee19b063285cb0bb73dea819f1be1ac5e82133`, and source manifest
`5db38f164ac6f461f1b92ae8f97f05419fccfada944ee55359b0056fe317aa1f`. Within
each role, exact and managed arms use the same retained executable; normal and
allocator roles use their separately retained normal and instrumented
executables. The pinned text is 10,000 bytes with SHA-256
`ad4fe690f0ef2281ad8e64a78d1f4d64e7c8625d672b3ac2f7e9fbc28a82f4af`; the
archive is 16,793,036 bytes with SHA-256
`a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4`.

All latency values are nanoseconds. Each comparison is same role, repeat, and
transport service model: exact reads versus the opt-in managed 4,096-byte
forward-start policy. The bootstrap interval is the deterministic 2,000
resample percentile 95% interval for that row's median. Percent changes are
`(managed - exact) / exact`; a comparison flag means the managed latency is
more than 5% higher at that percentile.

| role | repeat | transport | exact p50 / p95 / p99 | managed p50 / p95 / p99 | exact median CI | managed median CI | managed change p50 / p95 / p99 | latency flag |
| --- | ---: | --- | --- | --- | --- | --- | --- | --- |
| normal | 1 | 0 us | 386,977 / 398,672 / 421,842 | 390,782 / 402,692 / 407,492 | 384,641–389,236 | 388,466–394,487 | +0.983% / +1.008% / −3.402% | none |
| normal | 1 | 1,000 us + 100 MiB/s | 23,563,178 / 29,869,665 / 34,356,114 | 3,565,722 / 3,621,097 / 3,700,937 | 22,749,819–24,124,120 | 3,560,942–3,571,522 | −84.867% / −87.877% / −89.228% | none |
| normal | 2 | 0 us | 393,277 / 402,542 / 408,352 | 385,357 / 396,972 / 401,122 | 392,282–396,217 | 381,976–388,732 | −2.014% / −1.384% / −1.771% | none |
| normal | 2 | 1,000 us + 100 MiB/s | 20,466,051 / 20,491,261 / 20,496,651 | 3,649,518 / 3,749,988 / 5,847,349 | 20,461,561–20,477,141 | 3,635,687–3,657,788 | −82.168% / −81.700% / −71.472% | none |
| allocator | 1 | 0 us | 477,817 / 491,172 / 495,513 | 488,702 / 498,172 / 499,642 | 476,177–479,982 | 486,307–491,737 | +2.278% / +1.425% / +0.833% | none |
| allocator | 1 | 1,000 us + 100 MiB/s | 20,891,168 / 23,461,230 / 156,839,056 | 3,797,809 / 6,396,472 / 9,510,047 | 20,718,692–20,964,554 | 3,715,283–3,817,738 | −81.821% / −72.736% / −93.936% | none |
| allocator | 2 | 0 us | 489,667 / 500,523 / 501,063 | 488,083 / 497,702 / 519,822 | 484,772–493,517 | 484,662–490,927 | −0.323% / −0.564% / +3.744% | none |
| allocator | 2 | 1,000 us + 100 MiB/s | 21,737,072 / 24,948,573 / 25,649,997 | 3,655,943 / 3,761,049 / 3,765,658 | 21,451,246–22,438,061 | 3,651,968–3,673,248 | −83.181% / −84.925% / −85.319% | none |

The managed policy's raw source-read evidence is stable in all pilot and formal
rows: 19 logical requests, 16 hits, 3 misses, 3 fills, 3,966 logical requested
bytes, and 5,445 physical returned bytes. The exact arm has 19 physical calls
and 3,966 requested/returned bytes. The managed arm has 3 physical calls and
5,445 requested/returned bytes, with maximum physical request 4,096 bytes and
no short reads. Recomputing the eight pinned media ranges against every
physical trace gives zero overlap for exact and 69 requested/returned bytes
for managed. The 69 bytes come from the managed `[0, 4096)` fill; this is
measured synthetic transport overlap, not a filesystem or network observation.

Every row reports `open_successful_loads=0`, `successful_loads=1`,
`failed_loads=0`, and `budget_managed=true`. The managed and exact rows both
release memory and objects to zero after package/document drop, and each budget
input delta equals the physical returned-byte total. The candidate window is
configured at 4,096 bytes and is constructed inside the managed timing scope,
as required by the 0493 protocol.

The allocator executable reports the following p50/p95/p99 values. Allocation
calls, reallocation calls, and allocated bytes are constant across the 30
rows of each process; region peak is an absolute live-byte value, while the
increment is the separate operation peak increment.

| role | repeat | transport | allocation calls exact → managed | reallocation calls exact → managed | allocated bytes exact → managed | region peak p50 / p95 / p99 exact → managed | peak increment exact → managed |
| --- | ---: | --- | --- | --- | --- | --- | --- |
| allocator | 1 | 0 us | 5,327 → 5,330 | 128 → 128 | 1,755,981 → 1,760,365 | 19,131,826 / 19,138,846 / 19,139,366 → 19,130,730 / 19,132,566 / 19,132,702 | 139,171 → 143,555 |
| allocator | 1 | 1,000 us + 100 MiB/s | 5,327 → 5,330 | 128 → 128 | 1,755,981 → 1,760,365 | 19,132,031 / 19,139,051 / 19,139,571 → 19,130,935 / 19,132,771 / 19,132,907 | 139,171 → 143,555 |
| allocator | 2 | 0 us | 5,327 → 5,330 | 128 → 128 | 1,755,981 → 1,760,365 | 19,131,826 / 19,138,846 / 19,139,366 → 19,130,730 / 19,132,566 / 19,132,702 | 139,171 → 143,555 |
| allocator | 2 | 1,000 us + 100 MiB/s | 5,327 → 5,330 | 128 → 128 | 1,755,981 → 1,760,365 | 19,132,031 / 19,139,051 / 19,139,571 → 19,130,935 / 19,132,771 / 19,132,907 | 139,171 → 143,555 |

The managed window therefore adds three allocation calls and 4,384 allocated
bytes in this allocator build; reallocation calls remain equal and the peak
increment rises by 4,384 bytes. The absolute region peak is not itself the
increment and should not be interpreted as a standalone window cost.

GNU `time -v` records one whole-child maximum RSS observation per process,
including setup. The canonical analysis represents that singleton with equal
p50/p95/p99 values; it is not an operation-level or within-process RSS
distribution.

| role | repeat | transport | exact RSS KiB | managed RSS KiB |
| --- | ---: | --- | ---: | ---: |
| normal | 1 | 0 us | 94,952 | 95,008 |
| normal | 1 | 1,000 us + 100 MiB/s | 94,976 | 94,892 |
| normal | 2 | 0 us | 94,992 | 94,732 |
| normal | 2 | 1,000 us + 100 MiB/s | 94,812 | 95,080 |
| allocator | 1 | 0 us | 94,548 | 95,008 |
| allocator | 1 | 1,000 us + 100 MiB/s | 94,664 | 94,984 |
| allocator | 2 | 0 us | 94,520 | 94,888 |
| allocator | 2 | 1,000 us + 100 MiB/s | 93,400 | 93,480 |

The formal repeat-variance table below is separate from the exact-versus-
managed comparison. It compares repeat 1 with repeat 2 for each arm; a flag
means the absolute repeat change exceeds 5% at some latency percentile.

| role | arm | latency repeat 1 p50 / p95 / p99 → repeat 2 p50 / p95 / p99 | repeat change p50 / p95 / p99 | flag | RSS repeat 1 → repeat 2 KiB |
| --- | --- | --- | --- | --- | --- |
| normal | exact 0 us | 386,977 / 398,672 / 421,842 → 393,277 / 402,542 / 408,352 | +1.628% / +0.971% / −3.198% | none | 94,952 → 94,992 |
| normal | managed 0 us | 390,782 / 402,692 / 407,492 → 385,357 / 396,972 / 401,122 | −1.388% / −1.420% / −1.563% | none | 95,008 → 94,732 |
| normal | exact delayed | 23,563,178 / 29,869,665 / 34,356,114 → 20,466,051 / 20,491,261 / 20,496,651 | −13.144% / −31.398% / −40.341% | p50,p95,p99 | 94,976 → 94,812 |
| normal | managed delayed | 3,565,722 / 3,621,097 / 3,700,937 → 3,649,518 / 3,749,988 / 5,847,349 | +2.350% / +3.559% / +57.996% | p99 | 94,892 → 95,080 |
| allocator | exact 0 us | 477,817 / 491,172 / 495,513 → 489,667 / 500,523 / 501,063 | +2.480% / +1.904% / +1.120% | none | 94,548 → 94,520 |
| allocator | managed 0 us | 488,702 / 498,172 / 499,642 → 488,083 / 497,702 / 519,822 | −0.127% / −0.094% / +4.039% | none | 95,008 → 94,888 |
| allocator | exact delayed | 20,891,168 / 23,461,230 / 156,839,056 → 21,737,072 / 24,948,573 / 25,649,997 | +4.049% / +6.340% / −83.646% | p95 | 94,664 → 93,400 |
| allocator | managed delayed | 3,797,809 / 6,396,472 / 9,510,047 → 3,655,943 / 3,761,049 / 3,765,658 | −3.735% / −41.201% / −60.403% | p95,p99 | 94,984 → 93,480 |

The exact-versus-managed comparison has no latency or RSS adverse flag. The
only systematic positive adverse flags are physical requested and returned
bytes: every role, repeat, and transport arm records 5,445 versus 3,966,
which is +37.292% and reflects bounded forward overfetch. Allocation changes
are below 5%. Repeat-variance flags in the delayed arm describe run-to-run
transport timing variation; they do not establish a managed regression. No
result supports an overall 10× speedup or a filesystem/network claim. The
scope is a same-binary, same-role comparison over this synthetic managed
in-memory transport and pinned DOCX lifecycle, with descriptive evidence only.
