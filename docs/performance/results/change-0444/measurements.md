# OPC Part-addition observed baseline

One current revision; 12 reports, 360 retained samples, CPU 2, one worker,
30 samples and three warmups per report. Normal mode includes the ReadAt
observer and hashing sink. Allocator elapsed values include instrumentation.
These are not plain-source production latency estimates or speedup claims.

| Mode | Shape | Repeat | p50 ms | p95 ms | p99 ms | p50 bootstrap 95% ms | Whole-process peak RSS MiB |
| --- | --- | --- | ---: | ---: | ---: | --- | ---: |
| normal | tiny | R1 | 1.107870 | 1.113515 | 1.113585 | 1.105615–1.109905 | 80.805 |
| normal | medium | R1 | 8.017485 | 8.055815 | 8.108365 | 8.010796–8.031185 | 80.797 |
| normal | large | R1 | 61.875915 | 63.233921 | 63.495392 | 61.811764–61.969200 | 80.672 |
| allocator | tiny | R1 | 1.153865 | 1.163815 | 1.166725 | 1.151180–1.157915 | 80.680 |
| allocator | medium | R1 | 8.865229 | 8.893129 | 8.906500 | 8.858035–8.877004 | 80.656 |
| allocator | large | R1 | 65.234785 | 65.474590 | 65.475150 | 65.182309–65.303079 | 80.781 |
| allocator | large | R2 | 64.900679 | 65.078509 | 65.135130 | 64.870273–64.923703 | 80.797 |
| allocator | medium | R2 | 8.864569 | 8.905210 | 8.909120 | 8.848954–8.876769 | 80.770 |
| allocator | tiny | R2 | 1.155620 | 1.171445 | 1.180436 | 1.152380–1.158800 | 80.801 |
| normal | large | R2 | 61.688739 | 61.848356 | 61.896796 | 61.602634–61.715209 | 80.797 |
| normal | medium | R2 | 8.019791 | 8.054406 | 8.136496 | 8.010826–8.029565 | 80.781 |
| normal | tiny | R2 | 1.101234 | 1.111625 | 1.112405 | 1.100310–1.103335 | 80.797 |

p50 uses the midpoint of the central observations; p95/p99 use nearest rank.
Intervals use 2,000 deterministic bootstrap resamples within each invocation;
they do not measure machine-to-machine or day-to-day uncertainty. Raw samples,
all quantiles/intervals and every repeat check are retained in summary.json.

| Shape | Allocation calls | Requested bytes | Region peak above entry | Endpoint live delta | Source calls | Returned source bytes | Output bytes |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| tiny | 3,526 | 2,295,059 | 705,777 | 0 | 359 | 64,967 | 111,002 |
| medium | 51,538 | 7,464,851 | 2,716,017 | 0 | 5,162 | 1,008,481 | 779,088 |
| large | 205,145 | 24,040,435 | 9,181,553 | 0 | 20,531 | 4,027,537 | 2,916,848 |

Repeat review: **0 flags** above the frozen absolute 5% trigger.
Allocation values and source/sink counters match both repeats. No timed output
archive is retained; the endpoint live delta is zero because publication
consumes and drops its package. Preexisting input/oracle fixtures stay live.
Metadata/temporary peaks grow with Part count, so no bounded-total-memory
claim follows. Source calls include repeated reads and observer bookkeeping;
codec byte-flow, remote requests, lock wait and scaling are not measured.

## Profile review

The large normal whole-process profile reports **55.24% self time in the
instrumented source reader**, followed by SHA-256 at 9.79%. The reader scans
all ordinary member ranges for each read. This observer work can grow with
both read count and member count; the curve cannot establish production
topology complexity. A matched plain-source lifecycle is the next measurement
priority before a production optimization. No Amdahl speedup is inferred
from the whole-process fraction, which includes fixture construction, gates,
warmups, hashing and report work outside the timed interval.

Perf stat returned 11,538,555,261 cycles, 49,721,657,346 instructions,
4,964,600,258 branches and 8,676,980 branch misses. Its L1-miss event returned
zero on this guest; that is not evidence of zero hardware cache misses.
Perf report/script retained addr2line warnings. Symbol-level self rows are
usable, but precise source-line/inlined attribution remains limited. No
samples were reported lost. The raw profile, warnings and commands remain
in the bundle; profiles are diagnostic, not paired CPU improvement evidence.
