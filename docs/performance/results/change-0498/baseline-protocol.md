# Change 0498 initial source-backed Part baseline

This bundle is the pre-edit control for the bounded source-backed Part batch
work. It measures the existing `SourceBackedPackage` API. The normal row uses
the unmanaged compatibility constructor; the matched managed row opts into an
explicit `ExecutionContext`. The CSV label `managed-after-only` is retained as
a historical name from the pre-batch control protocol: managed source-backed
reads already existed before this change, and the label does not imply that
the capability was introduced by the batch API. Neither row is a production
batch or scaling claim.

The primary invocation is:

```text
RUSTUP_TOOLCHAIN=1.98.1 \
CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_BUILD_JOBS=2 \
taskset -c 0-7 /usr/bin/time -v \
/home/zhuhe/.cache/litchi-goal-0498/retained/source_backed_batch_perf-before \
  --corpus few-large --source owned --workers 1 \
  --warmups 3 --samples 30 --repeats 2 \
  --artifact-dir /tmp/litchi-change-0498-control \
  --output docs/performance/results/change-0498/baseline-owned-few-large-corrected.csv
```

The managed control is identical except for `--capability managed` and the
artifact/output names. Both CSVs retain all 60 measured samples and the five
percentile/throughput summary values. `/usr/bin/time -v` stderr is retained beside the CSVs and under
`development/captures` for whole-child RSS context; the CSV's cache and budget
columns report source-backed counters and release observations.

The corrected serial controls are:

| capability | p50 | p95 | p99 | mean | source calls | selected source bytes | retained bytes | memory after release |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| normal | 575,142 ns | 585,902 ns | 593,032 ns | 555,143 ns | 40 | 4,194,424 | 4,194,304 | 0 |
| managed-after-only | 580,352 ns | 591,442 ns | 596,042 ns | 570,191 ns | 40 | 4,194,424 | 4,194,304 | 0 |

The additional serial controls use the same 60-sample measured subset and are
retained in the per-case CSVs:

| corpus | source | capability | p50 | p95 | p99 |
| --- | --- | --- | ---: | ---: | ---: |
| few-large | FileSource | normal | 618,502 ns | 664,233 ns | 741,803 ns |
| few-large | short-read, 4 KiB, 25 µs/call | normal | 80,825,855 ns | 82,067,860 ns | 82,886,343 ns |
| few-large | FileSource | managed-after-only | 643,622 ns | 682,203 ns | 792,673 ns |
| few-large | short-read, 4 KiB, 25 µs/call | managed-after-only | 82,526,157 ns | 82,645,557 ns | 82,943,969 ns |
| many-small | owned | normal | 68,950 ns | 75,560 ns | 76,851 ns |
| many-small | owned | managed-after-only | 98,570 ns | 110,520 ns | 116,941 ns |
| many-small | FileSource | normal | 155,641 ns | 166,971 ns | 182,031 ns |
| many-small | short-read, 4 KiB, 25 µs/call | normal | 29,739,862 ns | 30,613,156 ns | 31,685,490 ns |
| many-small | FileSource | managed-after-only | 186,321 ns | 196,861 ns | 268,741 ns |
| many-small | short-read, 4 KiB, 25 µs/call | managed-after-only | 29,587,660 ns | 29,853,611 ns | 30,013,852 ns |

The measured interval excludes package open, ZIP catalog validation, URI
construction, corpus generation, file creation, and post-timer digest/byte
verification. It includes URI-to-Part lookup and each cold `PartView::data`
load, with all returned handles retained until after the cache/budget snapshot.
The current pre-edit executable calls `data_with_accounting` on the primary
serial path so the per-sample ZIP counters are available. Those counters are
not part of the future batch API contract; any post-edit comparison that uses
plain `data()` must treat these controls as accounting-instrumented controls,
or capture a separately built plain-data control before making a latency
comparison.
The source counters include the observer's positional calls during the timed
loads; the 40 calls include ZIP range metadata and four selected stored
payloads. The parent process was pinned to CPUs 0–7 on a shared virtualized
host, so these are descriptive controls rather than isolated-host evidence.

Each CSV contains 66 rows per case: warmup indexes 0–2 for each repeat and
measured indexes 3–32 for each repeat. Summaries use only rows with
`index >= 3`, yielding 60 measured samples. The retained `/usr/bin/time -v`
stderr receipts sit beside the CSVs and provide whole-child RSS context.

Two bounded `perf stat` control receipts are also retained for the future
matched comparison: `perfstat-baseline-managed-owned.txt` and
`perfstat-baseline-managed-short.txt`. They count cycles, instructions,
branches, branch misses, and cache misses for the whole child process,
including repeated package setup and post-timer verification. They do not
attribute counters to the timed Part-load interval.

The `baseline-owned-few-large.csv` file is an earlier exploratory run whose
digest was inside the timer and is retained only for traceability. The
`*-corrected.csv` files are the comparable controls used for review.
