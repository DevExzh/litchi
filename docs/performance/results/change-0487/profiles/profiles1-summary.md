# 0487 profile summary

Status: **complete**; profile validation: **pass**.

The 12 rows below are one whole diagnostic child each, with one sample and one warmup. The child includes setup and oracle work; profiler overhead is included. Missing PMU values remain unavailable. No elapsed-time or causal performance claim is made.

Profile attempt: `profiles1`. Before build: `f04a306ef192c835bd26d7b1585f8e8baedf0e0033f51e854893a5c038355ac2`. After build: `91f99a5d1a123c5b6e62c7b3179a5132ce7c6456ee7383c2da52cc1618bc7f86`.

## Child rows

| phase | tool | input | workload | receipt | report | RSS bytes | profile status |
|---|---|---|---|---|---|---:|---|
| before | perf | owned | s131072-a64-short-c64 | pass | pass | 75702272 | observed |
| before | perf | file | s131072-a64-short-c64 | pass | pass | 75636736 | observed |
| before | perf | owned | s64-a16384-short-c64 | pass | pass | 17846272 | observed |
| before | perf | file | s64-a16384-short-c64 | pass | pass | 17915904 | observed |
| after | perf | owned | s131072-a64-short-c64 | pass | pass | 75739136 | observed |
| after | strace | owned | s131072-a64-short-c64 | pass | pass | 75468800 | observed |
| after | perf | file | s131072-a64-short-c64 | pass | pass | 75292672 | observed |
| after | strace | file | s131072-a64-short-c64 | pass | pass | 75321344 | observed |
| after | perf | owned | s64-a16384-short-c64 | pass | pass | 17772544 | observed |
| after | strace | owned | s64-a16384-short-c64 | pass | pass | 18051072 | observed |
| after | perf | file | s64-a16384-short-c64 | pass | pass | 18116608 | observed |
| after | strace | file | s64-a16384-short-c64 | pass | pass | 17776640 | observed |

## Perf counters

Each row is a matched before/after whole-child observation. A review flag marks an absolute change of at least 5%; it is a review signal, not a performance claim.

| input | workload | event | before | before status | after | after status | percent change | review |
|---|---|---|---:|---|---:|---|---:|---|
| owned | s131072-a64-short-c64 | cycles | 7700041604 | measured | 7690675711 | measured | -0.12163431682154324 | False |
| owned | s131072-a64-short-c64 | instructions | 34584273205 | measured | 34543228279 | measured | -0.11868089798129965 | False |
| owned | s131072-a64-short-c64 | branches | 6529361927 | measured | 6518241468 | measured | -0.17031463601389668 | False |
| owned | s131072-a64-short-c64 | branch-misses | 9587713 | measured | 9474702 | measured | -1.1787065382537003 | False |
| owned | s131072-a64-short-c64 | page-faults | 23807 | measured | 22302 | measured | -6.321670097030285 | True |
| file | s131072-a64-short-c64 | cycles | 7666343100 | measured | 7690657355 | measured | 0.3171558418772048 | False |
| file | s131072-a64-short-c64 | instructions | 34614543390 | measured | 34561901048 | measured | -0.15208157278540949 | False |
| file | s131072-a64-short-c64 | branches | 6533064052 | measured | 6520432540 | measured | -0.19334743849837277 | False |
| file | s131072-a64-short-c64 | branch-misses | 9447338 | measured | 9633842 | measured | 1.9741434042055022 | False |
| file | s131072-a64-short-c64 | page-faults | 23815 | measured | 22303 | measured | -6.348939743858913 | True |
| owned | s64-a16384-short-c64 | cycles | 1860404645 | measured | 1734932220 | measured | -6.74436205785758 | True |
| owned | s64-a16384-short-c64 | instructions | 6223149178 | measured | 5898122542 | measured | -5.222864287891895 | True |
| owned | s64-a16384-short-c64 | branches | 1216413300 | measured | 1164199111 | measured | -4.292471070482376 | False |
| owned | s64-a16384-short-c64 | branch-misses | 2563711 | measured | 2149192 | measured | -16.168710123722995 | True |
| owned | s64-a16384-short-c64 | page-faults | 4081 | measured | 4092 | measured | 0.2695417789757412 | False |
| file | s64-a16384-short-c64 | cycles | 3028615234 | measured | 2242395060 | measured | -25.959724601979598 | True |
| file | s64-a16384-short-c64 | instructions | 8449804322 | measured | 6766649112 | measured | -19.91945784611508 | True |
| file | s64-a16384-short-c64 | branches | 1584081047 | measured | 1306811602 | measured | -17.503488569925427 | True |
| file | s64-a16384-short-c64 | branch-misses | 7801933 | measured | 4647718 | measured | -40.428634801144796 | True |
| file | s64-a16384-short-c64 | page-faults | 4079 | measured | 4083 | measured | 0.09806325079676391 | False |

## RSS review

GNU time contributes one maximum RSS observation per child; zero-baseline percentage changes are left undefined.

| input | workload | before | after | percent change | review |
|---|---|---:|---:|---:|---|
| owned | s131072-a64-short-c64 | 75702272 | 75739136 | 0.04869602856833676 | False |
| file | s131072-a64-short-c64 | 75636736 | 75292672 | -0.45489006823351025 | False |
| owned | s64-a16384-short-c64 | 17846272 | 17772544 | -0.41312829928850126 | False |
| file | s64-a16384-short-c64 | 17915904 | 18116608 | 1.1202560585276635 | False |

## After strace counts

The after strace rows are compared with retained 0485 after diagnostics only when the arm, prepared fixture identity, source identity, report identity, and trace filter bind exactly. Candidate archive framing differences are retained separately.

| input | workload | status | statx before/after | pread64 before/after |
|---|---|---|---:|---:|
| owned | s131072-a64-short-c64 | pass | 12 / 12 | 6 / 6 |
| file | s131072-a64-short-c64 | pass | 25219 / 21781 | 264 / 264 |
| owned | s64-a16384-short-c64 | pass | 12 / 12 | 6 / 6 |
| file | s64-a16384-short-c64 | pass | 1475055 / 591477 | 114 / 114 |

## Content identity

| pair | status | source | authored proof | candidate main XML | semantic identity | unchanged oracles | candidate archive |
|---|---|---|---|---|---|---|---|
| perf-owned-s131072-a64-short-c64 | pass | True | True | True | True | True | True |
| perf-file-s131072-a64-short-c64 | pass | True | True | True | True | True | True |
| perf-owned-s64-a16384-short-c64 | pass | True | True | True | True | True | True |
| perf-file-s64-a16384-short-c64 | pass | True | True | True | True | True | True |

## Custody

Source manifests differ in 2 path(s); the full path/hash diff is retained in JSON for review.

The summarizer binds the profile driver, sealed 0484 route validator hashes, the retained 0485 before build, the new 0487 after build, build gate receipts, binary hashes, report artifacts, resource artifacts, and profiler artifacts. It performs no capture or build work.
