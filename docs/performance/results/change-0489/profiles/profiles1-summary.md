# 0489 profile summary

Status: **complete**; profile validation: **pass**.

The 12 rows below are one whole diagnostic child each, with one sample and one warmup. The child includes setup and oracle work; profiler overhead is included. Missing PMU values remain unavailable. No elapsed-time or causal performance claim is made.

Profile attempt: `profiles1`. Before build: `91f99a5d1a123c5b6e62c7b3179a5132ce7c6456ee7383c2da52cc1618bc7f86`. After build: `74e37fb9f4da3c0964ed7e0e80c10414324953de31754f6d70176238b55d8afd`.

## Child rows

| phase | tool | input | workload | receipt | report | RSS bytes | profile status |
|---|---|---|---|---|---|---:|---|
| before | perf | owned | s131072-a64-short-c64 | pass | pass | 75468800 | observed |
| before | perf | file | s131072-a64-short-c64 | pass | pass | 75640832 | observed |
| before | perf | owned | s64-a16384-short-c64 | pass | pass | 17911808 | observed |
| before | perf | file | s64-a16384-short-c64 | pass | pass | 17600512 | observed |
| after | perf | owned | s131072-a64-short-c64 | pass | pass | 75505664 | observed |
| after | strace | owned | s131072-a64-short-c64 | pass | pass | 75571200 | observed |
| after | perf | file | s131072-a64-short-c64 | pass | pass | 75448320 | observed |
| after | strace | file | s131072-a64-short-c64 | pass | pass | 75923456 | observed |
| after | perf | owned | s64-a16384-short-c64 | pass | pass | 18149376 | observed |
| after | strace | owned | s64-a16384-short-c64 | pass | pass | 17711104 | observed |
| after | perf | file | s64-a16384-short-c64 | pass | pass | 17735680 | observed |
| after | strace | file | s64-a16384-short-c64 | pass | pass | 17866752 | observed |

## Perf counters

Each row is a matched before/after whole-child observation. A review flag marks an absolute change of at least 5%; it is a review signal, not a performance claim.

| input | workload | event | before | before status | after | after status | percent change | review |
|---|---|---|---:|---|---:|---|---:|---|
| owned | s131072-a64-short-c64 | cycles | 7673312519 | measured | 6573527427 | measured | -14.33259872156655 | True |
| owned | s131072-a64-short-c64 | instructions | 34590808093 | measured | 28886887204 | measured | -16.48970117629105 | True |
| owned | s131072-a64-short-c64 | branches | 6526097195 | measured | 5465384054 | measured | -16.25340704108223 | True |
| owned | s131072-a64-short-c64 | branch-misses | 9328107 | measured | 7667088 | measured | -17.806603204701663 | True |
| owned | s131072-a64-short-c64 | page-faults | 23815 | measured | 24570 | measured | 3.1702708377073274 | False |
| file | s131072-a64-short-c64 | cycles | 7713559115 | measured | 6800615009 | measured | -11.835575411934858 | True |
| file | s131072-a64-short-c64 | instructions | 34625219197 | measured | 28908488217 | measured | -16.510309862515786 | True |
| file | s131072-a64-short-c64 | branches | 6531327875 | measured | 5468203855 | measured | -16.277302875412605 | True |
| file | s131072-a64-short-c64 | branch-misses | 9528542 | measured | 8150833 | measured | -14.45875979766894 | True |
| file | s131072-a64-short-c64 | page-faults | 23812 | measured | 24567 | measured | 3.1706702502939694 | False |
| owned | s64-a16384-short-c64 | cycles | 1741637121 | measured | 1469191116 | measured | -15.643098192783638 | True |
| owned | s64-a16384-short-c64 | instructions | 5873454632 | measured | 4621173760 | measured | -21.321027409955143 | True |
| owned | s64-a16384-short-c64 | branches | 1159334705 | measured | 903632205 | measured | -22.055968729065174 | True |
| owned | s64-a16384-short-c64 | branch-misses | 2232756 | measured | 2203123 | measured | -1.3271938357796373 | False |
| owned | s64-a16384-short-c64 | page-faults | 4080 | measured | 4074 | measured | -0.14705882352941177 | False |
| file | s64-a16384-short-c64 | cycles | 2231962568 | measured | 1934171024 | measured | -13.342138809560895 | True |
| file | s64-a16384-short-c64 | instructions | 6765956888 | measured | 5513749615 | measured | -18.507467513145052 | True |
| file | s64-a16384-short-c64 | branches | 1306705894 | measured | 1051022210 | measured | -19.567041456996748 | True |
| file | s64-a16384-short-c64 | branch-misses | 4281213 | measured | 4227915 | measured | -1.2449275474030375 | False |
| file | s64-a16384-short-c64 | page-faults | 4083 | measured | 4072 | measured | -0.26940974773450893 | False |

## RSS review

GNU time contributes one maximum RSS observation per child; zero-baseline percentage changes are left undefined.

| input | workload | before | after | percent change | review |
|---|---|---:|---:|---:|---|
| owned | s131072-a64-short-c64 | 75468800 | 75505664 | 0.048846675712347354 | False |
| file | s131072-a64-short-c64 | 75640832 | 75448320 | -0.2545080413710944 | False |
| owned | s64-a16384-short-c64 | 17911808 | 18149376 | 1.3263206037045507 | False |
| file | s64-a16384-short-c64 | 17600512 | 17735680 | 0.7679776588317431 | False |

## After strace counts

The after strace rows are compared with retained 0487 after diagnostics only when the arm, prepared fixture identity, source identity, report identity, and trace filter bind exactly. Candidate archive framing differences are retained separately.

| input | workload | status | statx before/after | pread64 before/after |
|---|---|---|---:|---:|
| owned | s131072-a64-short-c64 | pass | 12 / 12 | 6 / 6 |
| file | s131072-a64-short-c64 | pass | 21781 / 21781 | 264 / 264 |
| owned | s64-a16384-short-c64 | pass | 12 / 12 | 6 / 6 |
| file | s64-a16384-short-c64 | pass | 591477 / 591477 | 114 / 114 |

## Content identity

| pair | status | source | authored proof | candidate main XML | semantic identity | unchanged oracles | candidate archive |
|---|---|---|---|---|---|---|---|
| perf-owned-s131072-a64-short-c64 | pass | True | True | True | True | True | True |
| perf-file-s131072-a64-short-c64 | pass | True | True | True | True | True | True |
| perf-owned-s64-a16384-short-c64 | pass | True | True | True | True | True | True |
| perf-file-s64-a16384-short-c64 | pass | True | True | True | True | True | True |

## Custody

Source manifests differ in 2 path(s); the full path/hash diff is retained in JSON for review.

The summarizer binds the profile driver, sealed 0484 route validator hashes, the retained 0487 before build, the new 0489 after build, build gate receipts, binary hashes, report artifacts, resource artifacts, and profiler artifacts. It performs no capture or build work.
