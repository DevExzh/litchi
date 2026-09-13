# change-0564 evidence packet: the shape of a source-backed XLS open

Change record: [`docs/performance/0564-xls-open-read-attribution.md`](../../0564-xls-open-read-attribution.md).
Disposition: attribution only. No production change, `performance_claim: none`.

## Contents

| Path | What it is |
| --- | --- |
| `traces/xls_file_source_open.pread.txt` | `strace -f -e trace=pread64` of one child with one warmup and one measured sample, so two opens. |
| `traces/xls_file_source_open.samples-1.strace.txt`, `...samples-11.strace.txt` | The `strace -f -c` isolation pair. Differencing and dividing by ten isolates one open. |
| `summarize_open_reads.py`, `open-reads.json` | The summarizer and its output. It reports only properties of the retained traces and measures no time. |

## Result

```
per open: 655 pread64, 636 statx
four-byte reads: 1242 of 1312 (94.7%)
reads continuing the previous read: 6 of 1311
size histogram: {'4': 1242, '512': 48, '840': 2, '1536': 2, '43473': 2, '49152': 2, '65536': 14}
```

## Replay

```sh
python3 -B docs/performance/results/change-0564/summarize_open_reads.py
```

To recapture from scratch, build the harness with the `xls-source-attribution`
feature and run:

```sh
strace -f -e trace=pread64 -o capture.pread.txt \
  tools/perf-baseline/target/release/xls_source_attribution \
  --input test-data/ole/xls/ConditionalFormattingSamples.xls \
  --mode file-source --operation open --warmups 1 --samples 1
```

Note that `--warmups 0` is rejected by the binary, so the isolation pair uses one
warmup with 1 and with 11 samples.

## What is not here

No production change, no latency or resource measurement, no cold-cache or
physical-device result. The per-group attribution in the record is a static
model validated against these totals, not an instrumented per-call-site capture.
