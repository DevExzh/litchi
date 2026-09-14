# Database-function performance: release 09

This report records one standalone candidate window for the OpenFormula
database functions. It contains 29 cases, three child processes per case (87
children), two warmups, and 31 measured iterations per child, pinned to CPU 6.
The release-07 archive is an earlier diagnostic window; release-09 is the
primary table below. Two common-value before/after windows and a separate RSS attribution probe
are recorded below. Peak-RSS review flags remain open, so these numbers do not
establish an unconditional performance acceptance decision.

## Custody and scope

The [release-09 receipt](release-09/receipt.json) binds the
[build archive](release-09/build.tar.gz) (SHA-256
`cdce42ba8ebe447cd8c43b6bfdb986675f74b0ba4698b7537119cc6feee97bf4`) and
[capture archive](release-09/capture.tar.gz) (SHA-256
`3f1c8df49f714096bd80fd6eb1cb4a2f3494c7c93f08451200a58aea316bb256`). The
capture's retained ELF is SHA-256
`cd8205d84958b8d395856254cc4a78e42489e8d3da26e4007a749430433c7ade`; the
captured runner is SHA-256
`ad2f49b8ef68a7d805b9190f3178a2d569383c216539376ea6f55188f62242eb`.
The candidate-09 source receipt has SHA-256
`67f44ff42fb3ee3321186a730a0ed77e65d89def7b1d72e4e87c08c299c13e52`, a
469-entry unchanged-source manifest, and source archive SHA-256
`f9d1cffba001d43b8046bf6dd79a9d44d4d90899267b4f743740a75d0be54a2c`.
The [source receipt](../candidate-09/receipt.json) is the source identity for
this window; its recorded Git head is context only.

The timer covers evaluator execution, instrumented resolver reads, checksum,
and result drop. Parsing, fixture construction, and the independent result
oracle are outside the timed region. `p50` is the child row's median elapsed
time normalized by its repeat count. `p95` is the child row's p95 similarly
normalized; both values in the table are then the median of the three rounds.
RSS is the maximum resident set size from external `/usr/bin/time -v`, shown
as median `[minimum,maximum]` across the rounds. It is not the execution
budget's memory counter.

`work/eval`, `alloc/free/eval`, `req=rel B/eval`, and provider counts are
normalized from the child batch. `peak live B` is the observed allocator peak
delta for that child batch, and `retained B` is the execution-budget
reservation while a result is live; neither is a transient-peak RSS claim.
Provider counts are from the instrumented in-memory resolver. The reads column
is `all/database/criteria/extent` per evaluation. All 29 cases have exact
requested/released-byte balance and live-before/live-after balance, and all
non-timing result fields match across the three rounds.

The reusable validator checks the outer receipt, build and source archive
hashes when adjacent provenance is present, all 261 child artifact hashes,
all 87 status/result pairings, and exact non-timing field sets and values:

```sh
python3 docs/report/spec-gap-validation-evidence/ods-formula-database-functions/performance/analyze_release.py \
  docs/report/spec-gap-validation-evidence/ods-formula-database-functions/performance/release-09/capture.tar.gz
```

Every raw child result, `/usr/bin/time` record, and stderr record is retained
inside the [capture archive](release-09/capture.tar.gz), named
`<round>-<case>.jsonl`, `.time`, and `.stderr`. The [build archive](release-09/build.tar.gz)
retains the exact harness and build receipt. The earlier
[release-07 receipt](release-07/receipt.json) remains a distinct source window
(candidate-07 source receipt SHA-256
`7795d72e7b4b6213e4724fae64c2918839c6f0ac86969f166bf1720b0f01365e`); later
changes preserving `Missing => NotAvailable` and removing the dead `Ignore`
criteria representation are outside that source-07 snapshot. No release-07
numbers are combined with the table below.

## Release-09 per-case results

| case | p50 ns/eval | p95 ns/eval | RSS KiB median [min,max] | work/eval | alloc/free/eval | req=rel B/eval | peak live B | retained B | reads all/db/crit/ext | expected |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---|
| database-daverage | 4736 | 4807 | 2900 [2844,2940] | 156 | 20/17 | 5880=5880 | 5288 | 0 | 19/17/2/2 | number |
| database-dcount | 4538 | 4613 | 2928 [2844,2940] | 154 | 20/17 | 5880=5880 | 5288 | 0 | 19/17/2/2 | number |
| database-dcounta | 4531 | 4576 | 2916 [2888,2940] | 155 | 20/17 | 5880=5880 | 5288 | 0 | 19/17/2/2 | number |
| database-dget | 4358 | 4413 | 2908 [2876,2972] | 157 | 20/17 | 5880=5880 | 5288 | 0 | 16/14/2/2 | number |
| database-dmax | 4513 | 4563 | 2908 [2876,2940] | 152 | 20/17 | 5880=5880 | 5288 | 0 | 19/17/2/2 | number |
| database-dmin | 4497 | 4532 | 2888 [2880,2936] | 152 | 20/17 | 5880=5880 | 5288 | 0 | 19/17/2/2 | number |
| database-dproduct | 4556 | 4611 | 2932 [2888,2940] | 156 | 20/17 | 5880=5880 | 5288 | 0 | 19/17/2/2 | number |
| database-dstdev | 4548 | 4620 | 2896 [2860,2940] | 154 | 20/17 | 5880=5880 | 5288 | 0 | 19/17/2/2 | number |
| database-dstdevp | 4538 | 4578 | 2920 [2888,2940] | 155 | 20/17 | 5880=5880 | 5288 | 0 | 19/17/2/2 | number |
| database-dsum | 4632 | 4683 | 2888 [2888,2940] | 152 | 20/17 | 5880=5880 | 5288 | 0 | 19/17/2/2 | number |
| database-dvar | 4551 | 4608 | 2876 [2876,2912] | 152 | 20/17 | 5880=5880 | 5288 | 0 | 19/17/2/2 | number |
| database-dvarp | 4502 | 4543 | 2888 [2880,2920] | 153 | 20/17 | 5880=5880 | 5288 | 0 | 19/17/2/2 | number |
| database-dcount-omitted-field | 4066 | 4105 | 2876 [2876,2936] | 131 | 20/17 | 5880=5880 | 5288 | 0 | 15/13/2/2 | number |
| database-dsum-reference-16 | 5445 | 5620 | 2900 [2896,2928] | 265 | 20/17 | 5880=5880 | 5288 | 0 | 34/32/2/2 | number |
| database-dsum-reference-256 | 26680 | 27090 | 2908 [2908,2936] | 3241 | 20/17 | 5880=5880 | 5288 | 0 | 418/416/2/2 | number |
| database-dsum-reference-4096 | 363571 | 370772 | 2900 [2872,3012] | 50857 | 20/17 | 5880=5880 | 5288 | 0 | 6562/6560/2/2 | number |
| database-dsum-wide-unused-4096 | 366862 | 372552 | 2936 [2932,2940] | 50935 | 20/17 | 6504=6504 | 5912 | 0 | 6588/6586/2/2 | number |
| database-dsum-criteria-rows-16 | 46895 | 49775 | 2924 [2844,2960] | 13926 | 20/17 | 6328=6328 | 5736 | 0 | 483/466/17/2 | number |
| database-dsum-criteria-columns-4-numeric-headers | 19345 | 19740 | 2888 [2876,3008] | 842 | 20/17 | 5960=5960 | 5352 | 0 | 270/262/8/2 | number |
| database-lazy-unselected-4096 | 640 | 700 | 2948 [2940,2948] | 14 | 4/4 | 632=632 | 632 | 0 | 0/0/0/0 | number |
| database-projected-cache-256 | 29910 | 30505 | 2932 [2876,2940] | 3297 | 37/34 | 7240=7240 | 6184 | 264 | 418/416/2/2 | array |
| database-dget-multiple-cardinality | 4530 | 4568 | 2928 [2840,2944] | 152 | 20/17 | 5880=5880 | 5288 | 0 | 19/17/2/2 | formula-error |
| database-empty-dsum | 4348 | 4430 | 2940 [2940,2940] | 212 | 20/17 | 5880=5880 | 5288 | 0 | 15/13/2/2 | number |
| database-empty-dmax | 4343 | 4385 | 2940 [2872,2940] | 212 | 20/17 | 5880=5880 | 5288 | 0 | 15/13/2/2 | number |
| database-empty-dmin | 4345 | 4392 | 2940 [2936,2980] | 212 | 20/17 | 5880=5880 | 5288 | 0 | 15/13/2/2 | number |
| database-empty-dproduct | 4322 | 4356 | 2940 [2852,2984] | 216 | 20/17 | 5880=5880 | 5288 | 0 | 15/13/2/2 | number |
| database-refusal-work | 295 | 295 | 2936 [2936,2940] | 0 | 4/4 | 288=288 | 288 | 0 | 0/0/0/0 | evaluation-failure |
| database-refusal-memory | 175 | 180 | 2888 [2812,2940] | 0 | 2/2 | 184=184 | 184 | 0 | 0/0/0/0 | evaluation-failure |
| database-refusal-cancelled | 15 | 20 | 2900 [2688,2932] | 0 | 0/0 | 0=0 | 0 | 0 | 0/0/0/0 | evaluation-failure |

The reference DSUM controls scale from 265 to 3,241 to 50,857 work units per
evaluation at sizes 16, 256, and 4096, with 34, 418, and 6,562 provider reads.
The wide-unused 4096 control has similar work and reads but a larger measured
request and live peak. The unselected lazy branch performs zero provider reads
and retains no result budget. The projected-cache case is the only successful
case retaining result budget in this window (264 bytes). Criteria-row and
criteria-column cases show different measured work and read counts by design;
they are separate corpus shapes, not a before/after comparison.

## Common-value comparison and follow-up

Both windows compare the retained `value-complex-09` baseline against
`value-database-09`, built from candidate-09 using the identical four-file
common-value harness. Each runs 27 cases in AB/BA/AB order, three warmups and
31 iterations per child, pinned to CPU 6: 162 children and 81 pairs per window.
[Common-09](common-09/receipt.json) retains the build and first capture;
[common-10](common-10/receipt.json) is a second capture of those same ELFs.
Both analyses report zero deterministic mismatches and zero instrumented
allocation/memory-counter changes. No production changes separate the windows.

The first window flagged one timing case and seven RSS cases at the 5% review
threshold. The second window did not reproduce the timing flag, but four RSS
cases remained above that threshold. These observations are retained separately;
the second window does not erase the first. Percentages below compare the
median of three child p50/p95/maximum-RSS values in each arm.

| Window | Flagged case | p50 delta | p95 delta | RSS delta |
| --- | --- | ---: | ---: | ---: |
| common-09 | reference-error | +1.23% | +0.59% | +9.66% |
| common-09 | reference-limit-cells | -0.21% | -0.21% | +9.09% |
| common-09 | reference-limit-memory | -1.26% | -1.30% | +8.86% |
| common-09 | reference-limit-work | -0.09% | -0.14% | +6.36% |
| common-09 | reference-range-1 | -0.84% | -0.90% | +7.80% |
| common-09 | reference-range-16 | -0.22% | -0.94% | +5.23% |
| common-09 | reference-range-256 | -3.64% | -3.08% | +7.27% |
| common-09 | reference-repeat-256 | +5.63% | +6.13% | -5.13% |
| common-10 | reference-limit-cells | +0.00% | +0.28% | +6.54% |
| common-10 | reference-limit-memory | -0.63% | -0.39% | +9.25% |
| common-10 | reference-range-16 | +0.07% | -0.49% | +6.72% |
| common-10 | reference-range-256 | -0.31% | -0.34% | +15.26% |

In common-10, `reference-repeat-256` was +0.19% at p50 and -0.15% at p95,
compared with +5.63%/+6.13% in common-09. The timing observation is therefore
not reproducible across these two windows; there is no basis here for a
speculative code optimization or a broad speedup claim.

The [RSS attribution probe](rss-09/receipt.json) ran six serial processes in
AB/BA/AB order, each evaluating the 256-cell range 25,000 times, and captured
three live `/proc/<pid>/smaps` snapshots per process. It is a sustained-loop
attribution experiment, not a replay of the short-lived peak-RSS benchmark.
The archive includes the probe, replayable analyzer, commands, raw stdout,
stderr, 18 snapshots, and hashes. Medians across nine snapshots in each arm:

| Resident mapping category | Baseline KiB | Database candidate KiB |
| --- | ---: | ---: |
| Executable mappings | 1372 | 1432 |
| Heap | 96 | 96 |
| Stack | 24 | 24 |
| Shared libraries | 2044 | 2044 |
| Other | 56 | 56 |
| Total | 3592 | 3644 |

Category medians are calculated independently and need not sum to the median
total. GNU `size` reports ELF text increasing from 1,757,912 to 1,858,288 bytes
(+100,376 bytes); that includes the new database implementation. The probe
supports a resident executable-mapping contribution, with unchanged median heap
and stack RSS in this scenario. It does not attribute every earlier peak-RSS
flag or establish that all workload memory costs are unchanged. Peak-RSS
acceptance remains open; measured allocator balance and source/result custody
are verified independently.

Reproduce either common analysis without extracting repository copies:

```sh
python3 docs/report/spec-gap-validation-evidence/ods-formula-matrix-functions/performance/analyze_common.py docs/report/spec-gap-validation-evidence/ods-formula-database-functions/performance/common-09/capture.tar.gz
python3 docs/report/spec-gap-validation-evidence/ods-formula-matrix-functions/performance/analyze_common.py docs/report/spec-gap-validation-evidence/ods-formula-database-functions/performance/common-10/capture.tar.gz
```

All captured archive members were compared byte-for-byte before loose scratch
was removed. The retained ELFs and one disk-backed workspace remain available
for further comparisons. No repository or build copies were added to tmpfs.
