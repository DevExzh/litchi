# Independent measurement review for change 0479

This review recomputes the measurements from the retained raw reports.  It read
the 24 `captures/r{1,2}-{normal,allocator}-{64,8192,131072}-{total,phases}.report.json`
files, their GNU `time -v` resource files, `protocol.json`,
`corpus-manifest.json`, and `summary.json`.  Each report has 30 measured samples
after three warmups, for 720 measured samples in total.  Repeat 2 used the
reverse capture order specified by the protocol.

Means below are arithmetic means.  Percentiles are nearest-rank percentiles
over the 30 raw values (`p50`, `p95`, and `p99`); elapsed values are converted
from nanoseconds to milliseconds.  RSS is one whole-process maximum per report
from GNU `time`, in KiB; it is not a 30-sample statistic.

## Corpus and protocol checks

The 24 reports contain the same corpus identity for each paragraph count and
match the frozen corpus manifest.  The independent raw checks also found the
same source reads and sink record between total and phases runs, phase
`ReadAt` counters summing to the lifecycle counters, and no allocator equation,
phase-boundary, or final-release failures in the measured allocator reports.
The source and candidate sizes are:

| paragraphs | source XML bytes | candidate XML bytes | source archive bytes | candidate archive bytes |
| ---: | ---: | ---: | ---: | ---: |
| 64 | 3,287 | 3,336 | 2,108 | 2,112 |
| 8,192 | 401,559 | 401,608 | 23,357 | 23,362 |
| 131,072 | 6,422,679 | 6,422,728 | 343,945 | 343,949 |

The candidate XML grows by exactly the copied 49-byte paragraph fragment in
each case.  The sink accepted exactly the candidate archive bytes and its hash
matched the candidate archive oracle in every report.  Total source I/O and
sink observations, constant across samples, were:

| paragraphs | source `ReadAt` calls | requested = returned bytes | sink accepted bytes | sink writes | largest write |
| ---: | ---: | ---: | ---: | ---: | ---: |
| 64 | 42 | 7,969 | 2,112 | 19 | 877 |
| 8,192 | 42 | 71,716 | 23,362 | 20 | 16,384 |
| 131,072 | 62 | 1,033,480 | 343,949 | 39 | 16,384 |

These I/O values were identical in normal and allocator binaries and in both
repeats.  The source request histogram included 20 requests in the
16,385–65,536 bucket at 131,072 paragraphs and no request over 65,536; this is
an observed shape, not a source adapter limit.

## Normal-binary latency and external RSS

The primary total-lifecycle observations are:

| paragraphs | repeat | mean ms | p50 ms | p95 ms | p99 ms | GNU time max RSS KiB |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 64 | 1 | 0.169641 | 0.167321 | 0.183771 | 0.185671 | 4,548 |
| 64 | 2 | 0.169123 | 0.167691 | 0.176081 | 0.180731 | 4,788 |
| 8,192 | 1 | 17.011582 | 17.011344 | 17.127055 | 17.141115 | 10,924 |
| 8,192 | 2 | 16.251897 | 16.235436 | 16.346136 | 16.607368 | 11,180 |
| 131,072 | 1 | 271.346097 | 271.104147 | 274.986102 | 275.231225 | 113,448 |
| 131,072 | 2 | 271.230824 | 270.497808 | 278.344301 | 278.702622 | 113,388 |

The phases runs are separate executions.  Summing the six phase elapsed values
within each sample gives the following attribution-oriented distribution; this
sum is not a replacement for the total-mode timing.

| paragraphs | repeat | phase-sum mean ms | p50 ms | p95 ms | p99 ms | GNU time max RSS KiB |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 64 | 1 | 0.172300 | 0.169680 | 0.181901 | 0.203730 | 4,744 |
| 64 | 2 | 0.168985 | 0.168052 | 0.175701 | 0.176800 | 4,780 |
| 8,192 | 1 | 16.809483 | 16.826722 | 16.956735 | 17.007385 | 11,180 |
| 8,192 | 2 | 16.456190 | 16.439816 | 16.601228 | 16.609788 | 11,180 |
| 131,072 | 1 | 273.519877 | 273.987937 | 276.849901 | 277.969555 | 113,644 |
| 131,072 | 2 | 274.460829 | 274.280513 | 278.412741 | 278.500391 | 113,540 |

For comparison, allocator-binary total latency was 0.206283/0.206873 ms at
64, 20.649763/20.503646 ms at 8,192, and 337.422280/336.273746 ms at
131,072 for repeats 1/2 respectively.  Its total-mode p50/p95/p99 values were
0.203301/0.216711/0.241842 ms, 20.660210/20.728220/20.735490 ms, and
337.153031/339.922003/340.866887 ms for repeat 1.  The corresponding repeat-1
external RSS values were 4,748, 10,972, and 113,328 KiB.

## Allocator accounting

For each total-mode allocator sample, the operation incremental peak is

`region_peak_live_bytes - live_bytes_before`.

The table reports repeat 1 distributions; every value in these allocation
fields was identical across the 30 samples and across repeat 2.  `alloc bytes`
is the harness `allocated_bytes` counter, and `alloc/realloc calls` keeps the
ordinary allocation and reallocation callback counts separate.

| paragraphs | incremental peak bytes | allocated bytes | alloc calls | realloc calls | total retention (`after - before`) |
| ---: | ---: | ---: | ---: | ---: | ---: |
| 64 | 509,974 | 1,036,969 | 1,888 | 67 | 0 |
| 8,192 | 3,071,310 | 1,615,057,345 | 213,346 | 16,453 | 0 |
| 131,072 | 41,793,870 | 549,272,923,585 | 3,653,986 | 507,973 | 0 |

The raw absolute region peaks were 534,199, 3,138,041, and 42,501,779 bytes
on average for repeat 1.  They are reported here only to make clear what was
subtracted; the operation comparison uses the incremental values above because
the process allocator already had live bytes at region entry.

Phase attribution from repeat 1 is below.  `reads` is `calls/requested
bytes`; returned bytes equal requested bytes for all rows.  `retained` is
`live_bytes_after - live_bytes_before`.  `peak-before` is the phase region
high-water increment.  Allocation bytes include the full requested `new_size`
on a reallocation callback.

| paragraphs | phase | normal mean ms | allocator mean ms | reads | retained bytes | peak-before bytes | alloc bytes | alloc/realloc calls |
| ---: | :--- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 64 | open | 0.015277 | 0.016772 | 13/893 | +3,408 | 132,977 | 306,689 | 151/14 |
| 64 | snapshot | 0.032728 | 0.041537 | 6/2,493 | +4,748 | 83,896 | 153,780 | 433/15 |
| 64 | stage | 0.026791 | 0.035356 | 0/0 | +4,432 | 5,979 | 6,730 | 414/8 |
| 64 | commit | 0.000051 | 0.000049 | 0/0 | 0 | 0 | 0 | 0/0 |
| 64 | publish | 0.097206 | 0.113211 | 23/4,583 | +1,844 | 497,386 | 569,770 | 890/30 |
| 64 | drop | 0.000247 | 0.000384 | 0/0 | -14,432 | 0 | 0 | 0/0 |
| 8,192 | open | 0.020590 | 0.022142 | 13/893 | +3,408 | 132,977 | 306,689 | 151/14 |
| 8,192 | snapshot | 3.491814 | 4.461071 | 6/44,991 | +533,068 | 664,687 | 403,481,316 | 53,297/4,111 |
| 8,192 | stage | 3.384047 | 4.367004 | 0/0 | +532,752 | 664,347 | 403,465,354 | 53,279/4,105 |
| 8,192 | commit | 0.000052 | 0.000052 | 0/0 | 0 | 0 | 0 | 0/0 |
| 8,192 | publish | 9.889338 | 11.792551 | 23/25,832 | +660,212 | 2,002,082 | 807,803,986 | 106,619/8,223 |
| 8,192 | drop | 0.023641 | 0.024437 | 0/0 | -1,729,440 | 0 | 0 | 0/0 |
| 131,072 | open | 0.072526 | 0.074488 | 13/893 | +3,408 | 132,977 | 306,689 | 151/14 |
| 131,072 | snapshot | 56.870787 | 72.614541 | 21/686,167 | +8,520,268 | 10,617,967 | 137,315,271,396 | 913,457/126,991 |
| 131,072 | stage | 56.062721 | 72.613474 | 0/0 | +8,519,952 | 10,617,627 | 137,317,221,514 | 913,439/126,985 |
| 131,072 | commit | 0.000151 | 0.000249 | 0/0 | 0 | 0 | 0 | 0/0 |
| 131,072 | publish | 159.679907 | 191.676630 | 28/346,420 | +10,613,492 | 24,750,242 | 274,640,123,986 | 1,826,939/253,983 |
| 131,072 | drop | 0.833785 | 0.900677 | 0/0 | -27,657,120 | 0 | 0 | 0/0 |

The six allocator phase retained deltas close back to zero in every sample;
adjacent phase `live_bytes_after` and `live_bytes_before` values were
continuous.  Phase high-water increments are independent attribution regions
and must not be summed into the total-operation peak.

## Repeat review

The repeat comparison uses `(repeat 2 - repeat 1) / repeat 1`, with an absolute
review flag at strictly greater than 5%.  There were no greater-than-5% flags
for total latency in any allocator lane, or for normal total latency.  The
following are the flags retained by the raw comparison; small phase durations
make their percentile changes especially quantized:

| lane | fields over 5% from repeat 1 to repeat 2 |
| :--- | :--- |
| normal / 64 / total | whole-process RSS +5.28% |
| normal / 64 / phases | open p99 +9.50%; snapshot p99 -8.97%; stage p99 -24.45%; commit p99 +16.67%; publish p99 -20.02%; drop p95 +10.34%; drop p99 +13.79% |
| normal / 8,192 / phases | open p95 -24.21%, p99 -24.51%; commit mean +23.23%, p50 +20.00%, p95 +50.00%, p99 +50.00%; drop p99 -14.63% |
| normal / 131,072 / phases | open p95 -6.52%, p99 -7.14%; commit mean -14.16%, p50 -23.08%, p95 +14.81%, p99 +46.67%; drop p50 +5.58% |
| allocator / 64 / phases | snapshot p95 -14.39%, p99 -36.01%; stage p99 +18.21%; commit p95 +20.00%, p99 -14.29%; publish p99 +8.02% |
| allocator / 8,192 / phases | open p95 -18.55%, p99 -17.64%; commit mean +19.75%, p95 +66.67%, p99 +42.86%; drop p95 -7.66%, p99 -32.30% |
| allocator / 131,072 / phases | commit mean +11.91%, p99 -8.93%; drop p95 -8.93% |

These flags are review signals between two process repeats, not claims of a
latency regression.  In particular, the total-mode means that dominate the
large cases were stable: normal 131,072 changed by -0.04% and allocator
131,072 by -0.34%.

## Summary comparison and limits

The current `summary.json` is schema
`docx-plain-paragraph-tail-append-summary-v1` with 24 rows and 720 samples.  I
recomputed the latency percentiles, external RSS, allocator incremental peaks,
phase retained deltas, and their raw report identities independently; all
compared values matched the summary rows.  Its recorded corpus-manifest hash
matches the current manifest (`880812c9467453b38eeabfcb5680b9fd1f0c9be39a19ef0bfee9c5cb9ae630b8`).

The counting allocator's `allocated_bytes` includes the full requested
`new_size` for reallocations.  There is no physical-copy-byte counter, so these
measurements do not estimate bytes copied by reallocations.  The normal binary
does not expose allocator metrics; its allocation fields are therefore
unavailable, not zero.

`process.rss_bytes` in the per-sample observer is a saturating RSS delta, while
`process.peak_rss_bytes` is an absolute VmHWM endpoint.  The GNU `time` value
reported above has broader whole-child scope: setup, corpus and oracle
construction, warmups, measured lifecycles, report serialization, and teardown
all contribute to that maximum.  Corpus and independent source/candidate
oracles are built outside each measured lifecycle, even though they remain in
the external process RSS envelope.

The total run includes publication, sink digest finalization, and owner drops.
The phases run expose the same lifecycle as six separate non-nested attribution
regions.  The harness exercises one plain-paragraph copy per lifecycle at the
tail.  It does not exercise repeated 64/256 operation reopen cycles or an
explicit bounded-window append mechanism, so these observations do not prove a
constant-memory or unbounded-append property.  Any production optimization
decision remains conditional on a separately scoped implementation and profile
review.
