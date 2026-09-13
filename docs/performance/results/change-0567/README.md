# change-0567 evidence packet: one archive index per OOXML open

Change record:
[`docs/performance/0567-ooxml-single-index-per-open.md`](../../0567-ooxml-single-index-per-open.md).
Disposition: attribution only. No production change, `performance_claim: none`.

## Contents

| Path | What it is |
| --- | --- |
| `segment_by_process.py`, `construction-attribution.json` | Re-derives the correction to change 0561 from the traces change 0562 already retained, by counting end-of-central-directory reads per process instead of per child. |
| `traces/loader-840-at-64.txt` | The resolved stack for the read change 0562 called "a structural member being re-read", captured freshly on the change-0565 baseline binary. |
| `counts.json`, `counts-full.json` | Per-scenario construction counts and positional reads, for two feature sets. |
| `count.py`, `members.py`, `analyze.py`, `attr.py` | The counters. `count.py` is descriptor-aware; `members.py` maps each 30-byte local-header read to a member using the fixture's own central directory. |

## Replay

The correction to change 0561 replays entirely from evidence already in the
repository:

```sh
python3 -B docs/performance/results/change-0567/segment_by_process.py
```

It reports, for all four of change 0562's retained captures:

```
docx_file_source_full_text     total 10  processes 5  per-process [2]
docx_file_source_open          total 10  processes 5  per-process [2]
pptx_file_source_open          total 10  processes 5  per-process [2]
pptx_file_source_selected_slide total 10 processes 5  per-process [2]
```

Ten constructions per traced child is **two per process across five processes**,
which is the lifecycle change 0561 itself describes: one untimed preparing open
and one post-timer oracle open. One library-level open builds the index once.

The correction to change 0562 was recaptured, because the original stack-resolved
trace was not retained:

```sh
strace -f -k -e trace=pread64,openat -o loader.txt \
  taskset -c 17 <baseline litchi-perf-baseline> \
    --warmup 0 --samples 1 --case docx_file_source_open --json /dev/null
```

The 840-byte read at offset 64 resolves to the dynamic loader reading libc's ELF
program-header table before `main`, on a descriptor opened for
`/usr/lib/x86_64-linux-gnu/libc.so.6` that the package file later reuses. It is
not attributable to this library.

## Why descriptor-aware counting is necessary

Plain process-and-descriptor segmentation is unsafe on these traces, which is how
the original reading arose. The dynamic loader uses descriptor 3 before `main`,
and the package file later reuses descriptor 3. Attribution has to follow
`openat` and `close` and count only reads on the descriptor that actually holds
the package.

## What is not here

The probe packages and their 5.2 GiB of build output were removed after capture.
No cold-cache, latency-bearing-source, allocation or multi-fixture-corpus result
is claimed. The wall-clock medians quoted in the record are paired microbenchmark
figures from those probes, not retained here; the construction counts, which are
the record's load-bearing result, are retained and replayable above.
