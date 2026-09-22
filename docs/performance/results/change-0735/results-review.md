# 0735 PPT record-staging results review

This independent review reads the retained raw captures and the derived
`analysis.json`/`audit.json`. It does not run Cargo, a native probe, or a
profiler, and it does not alter the frozen capture, source, or build inputs.

Disposition: **reject the candidate production hunk and restore the archived
baseline**. The direct staging transition does reduce measured allocation
work, but the public latency result is fixture-sensitive in the wrong way: the
independent secondary deck has a repeatable p50 regression above the packet's
5% review threshold, while its mean is positive in every paired process. The
primary p50 improvement is smaller than that secondary regression, and peak
live bytes do not change on either fixture.

## Raw evidence check

The capture manifest contains the fixed 48-process matrix: 36 native processes
and 12 allocation processes, with CPU 12 pinning, serialized execution, three
cycles, three repeats, two variants, and both fixtures. Native reports contain
50 samples and three warmups; allocation reports contain one sample and no
warmup. The 18 native baseline/candidate process pairs and six allocation
baseline/candidate repeat sets were recomputed directly from the raw JSON.

The independent custody scan found no missing or changed output/stderr digest,
sample-count or schedule mismatch, false nested oracle witness, or malformed
rejection control. Every report retained all eight corruption controls as
rejected. The derived report and independent audit both report `passed`, and
the audit's `analysis_match` is true. No selective rerun or discarded tail was
used.

The bound evidence identities are:

| artifact | SHA-256 |
| --- | --- |
| [`analysis.json`](analysis.json) | `dbacd3b3c2b13e222d66ea8a49adb5bb1bb91e3ac154e5aa9bffc5ce9bdd7def` |
| [`audit.json`](audit.json) | `9b85968109230d4103ee453ff7f4dc65ed5b3afcc6201153dfae5c1ea38bc7ec` |
| [`captures/manifest.json`](captures/manifest.json) | `16a33514b82a3a4cb725fa7009c0938110185cd2d4fe3fda13c784677448d12c` |
| [`freeze.json`](freeze.json) | `68c0a25274703f4b971576dce4c953187ed11db29e6eba8ab979acb33c3f95c9` |

## Latency result

The percentages below are candidate minus baseline, divided by the baseline,
computed per process pair. The interval is the deterministic nine-pair
bootstrap percentile interval for the median pair percentage, using 10,000
resamples and seed 7335.

| fixture | p50 paired median | p50 95% interval | mean paired median | p50 pairs beyond 5% |
| --- | ---: | ---: | ---: | ---: |
| primary `45543.ppt` | −3.5299% | [−3.7524%, −3.1449%] | −5.2900% | 0/9 |
| secondary `41246-1.ppt` | +6.3363% | [+5.4023%, +6.6149%] | +1.8466% | 8/9 |

The primary p50 improvement is present in all nine pairs but remains below the
5% p50 review threshold. Its mean, p95, p99, and maximum flags are retained in
the raw comparison; p95, p99, and maximum exceed 5% in all nine primary pairs,
and mean exceeds 5% in eight. On the secondary, p50 is the only timing field
with a review flag: eight pairs exceed +5%, with one +3.7728% pair. The
secondary mean is positive in all nine pairs even though it remains below the
5% field threshold. These tails and field-level flags are part of the
disposition, not grounds for a selective rerun.

## Allocation result

The allocation lane supports the intended mechanism but does not offset the
latency decision:

| fixture | allocated bytes | allocation calls | peak live bytes | retained bytes |
| --- | ---: | ---: | ---: | ---: |
| primary | −696,415 (−6.1133%) | −27 (−0.4772%) | 0 | 0 |
| secondary | −331,274 (−3.1717%) | −32 (−0.1715%) | 0 | 0 |

Deallocated bytes fall by the same absolute amount as allocated bytes in each
fixture. The process-local peak-live and retained ownership counters are
unchanged. The allocation lane therefore shows less clone-related byte work,
but no reduction in the measured live-memory boundary and no latency benefit
on the secondary case.

The matrix measures the public PPT format open/edit/commit workflow on two
fixtures. It does not measure insertion throughput, multi-record scaling,
cold I/O, RSS, instructions, concurrency, or broad CRUD behavior. Those
limits remain in force after rejecting this candidate.
