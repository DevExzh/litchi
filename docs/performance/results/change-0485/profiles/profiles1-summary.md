# 0485 profile summary

Status: **complete**; profile validation: **pass**.

The 12 rows below are one whole diagnostic child each, with one sample and one warmup. The child includes setup and oracle work; profiler overhead is included. Missing PMU values remain unavailable. No elapsed-time or causal performance claim is made.

Profile attempt: `profiles1`. Before build: `6c245a5fad0087b2218a3e3252e7e1350b256b05235491d7241bf2d47b7d6d1e`. After build: `f04a306ef192c835bd26d7b1585f8e8baedf0e0033f51e854893a5c038355ac2`.

## Child rows

| phase | tool | input | workload | receipt | report | RSS bytes | profile status |
|---|---|---|---|---|---|---:|---|
| before | perf | owned | s131072-a64-short-c64 | pass | pass | 75341824 | observed |
| before | perf | file | s131072-a64-short-c64 | pass | pass | 75558912 | observed |
| before | perf | owned | s64-a16384-short-c64 | pass | pass | 17907712 | observed |
| before | perf | file | s64-a16384-short-c64 | pass | pass | 17551360 | observed |
| after | perf | owned | s131072-a64-short-c64 | pass | pass | 75702272 | observed |
| after | strace | owned | s131072-a64-short-c64 | pass | pass | 75472896 | observed |
| after | perf | file | s131072-a64-short-c64 | pass | pass | 75939840 | observed |
| after | strace | file | s131072-a64-short-c64 | pass | pass | 75493376 | observed |
| after | perf | owned | s64-a16384-short-c64 | pass | pass | 17715200 | observed |
| after | strace | owned | s64-a16384-short-c64 | pass | pass | 18239488 | observed |
| after | perf | file | s64-a16384-short-c64 | pass | pass | 17752064 | observed |
| after | strace | file | s64-a16384-short-c64 | pass | pass | 17981440 | observed |

## Perf counters

Each row is a matched before/after whole-child observation. A review flag marks an absolute change of at least 5%; it is a review signal, not a performance claim.

| input | workload | event | before | before status | after | after status | percent change | review |
|---|---|---|---:|---|---:|---|---:|---|
| owned | s131072-a64-short-c64 | cycles | 8873401392 | measured | 7634401351 | measured | -13.963078939684237 | True |
| owned | s131072-a64-short-c64 | instructions | 41013866960 | measured | 34504179985 | measured | -15.871917128294113 | True |
| owned | s131072-a64-short-c64 | branches | 8037147949 | measured | 6514658067 | measured | -18.943161077300207 | True |
| owned | s131072-a64-short-c64 | branch-misses | 10192956 | measured | 9714498 | measured | -4.694006331431235 | False |
| owned | s131072-a64-short-c64 | page-faults | 23819 | measured | 22293 | measured | -6.406650153239011 | True |
| file | s131072-a64-short-c64 | cycles | 8872661718 | measured | 7681689246 | measured | -13.422944656887685 | True |
| file | s131072-a64-short-c64 | instructions | 41036513708 | measured | 34613093019 | measured | -15.652939561841396 | True |
| file | s131072-a64-short-c64 | branches | 8040382273 | measured | 6532752590 | measured | -18.750721443465377 | True |
| file | s131072-a64-short-c64 | branch-misses | 10540088 | measured | 9878810 | measured | -6.273932437755738 | True |
| file | s131072-a64-short-c64 | page-faults | 23819 | measured | 23808 | measured | -0.04618161971535329 | False |
| owned | s64-a16384-short-c64 | cycles | 1994928298 | measured | 1887139792 | measured | -5.40312682456119 | True |
| owned | s64-a16384-short-c64 | instructions | 6958340532 | measured | 6223168664 | measured | -10.56533328052994 | True |
| owned | s64-a16384-short-c64 | branches | 1389064951 | measured | 1216389812 | measured | -12.431034191431413 | True |
| owned | s64-a16384-short-c64 | branch-misses | 2863765 | measured | 2589299 | measured | -9.58409646042884 | True |
| owned | s64-a16384-short-c64 | page-faults | 4087 | measured | 4073 | measured | -0.342549547345241 | False |
| file | s64-a16384-short-c64 | cycles | 4914001216 | measured | 3068311268 | measured | -37.55981870721621 | True |
| file | s64-a16384-short-c64 | instructions | 12596868904 | measured | 8467125935 | measured | -32.783884634130345 | True |
| file | s64-a16384-short-c64 | branches | 2320113439 | measured | 1587458529 | measured | -31.578408955545903 | True |
| file | s64-a16384-short-c64 | branch-misses | 16041474 | measured | 8133136 | measured | -49.29932249368107 | True |
| file | s64-a16384-short-c64 | page-faults | 4086 | measured | 4081 | measured | -0.12236906510034264 | False |

## RSS review

GNU time contributes one maximum RSS observation per child; zero-baseline percentage changes are left undefined.

| input | workload | before | after | percent change | review |
|---|---|---:|---:|---:|---|
| owned | s131072-a64-short-c64 | 75341824 | 75702272 | 0.47841687506795694 | False |
| file | s131072-a64-short-c64 | 75558912 | 75939840 | 0.5041470157749227 | False |
| owned | s64-a16384-short-c64 | 17907712 | 17715200 | -1.0750228728270814 | False |
| file | s64-a16384-short-c64 | 17551360 | 17752064 | 1.1435239206534422 | False |

## After strace counts

The after strace rows are compared with retained 0484 metadata2 summaries only when the arm, prepared fixture identity, source identity, report identity, and trace filter bind exactly. Candidate archive framing differences are retained separately.

| input | workload | status | statx before/after | pread64 before/after |
|---|---|---|---:|---:|
| owned | s131072-a64-short-c64 | pass | 12 / 12 | 6 / 6 |
| file | s131072-a64-short-c64 | pass | 15415 / 25219 | 264 / 264 |
| owned | s64-a16384-short-c64 | pass | 12 / 12 | 6 / 6 |
| file | s64-a16384-short-c64 | pass | 3735939 / 1475055 | 114 / 114 |

## Content identity

| pair | status | source | authored proof | candidate main XML | semantic identity | unchanged oracles | candidate archive |
|---|---|---|---|---|---|---|---|
| perf-owned-s131072-a64-short-c64 | pass | True | True | True | True | True | True |
| perf-file-s131072-a64-short-c64 | pass | True | True | True | True | True | True |
| perf-owned-s64-a16384-short-c64 | pass | True | True | True | True | True | True |
| perf-file-s64-a16384-short-c64 | pass | True | True | True | True | True | True |

## Custody

Source manifests differ in 1 path(s); the full path/hash diff is retained in JSON for review.

The summarizer binds the profile driver, sealed 0484 route validator hashes, both normal build records, build gate receipts, binary hashes, report artifacts, resource artifacts, and profiler artifacts. It performs no capture or build work.
