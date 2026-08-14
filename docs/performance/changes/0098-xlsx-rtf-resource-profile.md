# Change 0098: XLSX/RTF resource-profile evidence

Date: 2026-08-14

Status: Recorded evidence; no production code change.

## Scope

This record adds external allocation and peak-memory evidence for the frozen
XLSX source-provenance and RTF bounded-ASCII binaries. It covers the matched
XLSX source-backed exact-256 selector on the `medium` and `dense-sparse`
corpora, and fresh streaming RTF creation on the `medium` and `large` corpora.
The binaries were invoked directly; Cargo was not part of any profile.

The before binary is
`/tmp/litchi-perf-baseline-before-xlsx-managed-0108` (SHA-256
`04326ec7296a8c9270cfbfc19b3789966a2f89bc70f381def7dd438fc6a56d43`). The
after binary is `/tmp/litchi-perf-baseline-after-xlsx-rtf-0108` (SHA-256
`0b604c3872bc9c559f39bff2916525e5efbe8d19939f20f28e37b58277c2d8a5`). Both
reports embed revision `7fc9ce27c2b5d42cb5afd8900cfd6222661ae261` and report a
dirty worktree; the executable hashes are the authoritative identity of the
frozen inputs. The associated production changes are `85ec86106` (XLSX) and
`d38cd455d` (RTF).

The harness reports Linux 6.8.0-101-generic on an AMD EPYC 9575F 64-Core
Processor, Rust `1.95.0`, the Rust system allocator, 33,605,935,104 bytes
reported memory, 4 KiB pages, and CPU affinity 2. The harness reports one logical CPU
available inside the execution environment. Heaptrack was 1.5.0, GNU
`/usr/bin/time` was available, and `perf` was 6.8.12.

## Method

For each of the four cases, five independent process invocations used CPU 2,
five untimed warm-ups, and 30 measured samples. The process-level RSS and
user/wall-clock values below come from `/usr/bin/time -v`; the p50/p95/p99
values are the harness's per-process 30-sample summaries. Corpus generation,
validation, and the complete sink checks remain in the process, as in the
normal harness contract.

Heaptrack used one warm-up and one measured sample per case to keep the
compressed captures bounded. Its allocation bytes are the sum of
`allocation_size * allocation_count` from `heaptrack_print -H`'s histogram;
allocation calls and temporary allocations are process totals. Heaptrack peak
RSS includes heaptrack overhead. These are whole-process profiles including
startup, deterministic corpus construction, one warm-up, and one measured
sample; they are not per-operation allocator attribution.

The exact command families (with `rN` expanded from 1 through 5) were:

```sh
/usr/bin/time -v taskset -c 2 /tmp/litchi-perf-baseline-before-xlsx-managed-0108 \
  --warmup 5 --samples 30 \
  --case xlsx_source_backed_cell_values_batch_edit_save \
  --xlsx-cell-crud-shape medium \
  --json /tmp/resprofile-before-xlsx-medium-0109-rN.json

/usr/bin/time -v taskset -c 2 /tmp/litchi-perf-baseline-after-xlsx-rtf-0108 \
  --warmup 5 --samples 30 \
  --case xlsx_source_backed_cell_values_batch_edit_save \
  --xlsx-cell-crud-shape dense-sparse \
  --json /tmp/resprofile-after-xlsx-dense-sparse-0109-rN.json

/usr/bin/time -v taskset -c 2 /tmp/litchi-perf-baseline-before-xlsx-managed-0108 \
  --warmup 5 --samples 30 --case rtf_streaming_create \
  --semantic-shape medium \
  --json /tmp/resprofile-before-rtf-medium-0109-rN.json

/usr/bin/time -v taskset -c 2 /tmp/litchi-perf-baseline-after-xlsx-rtf-0108 \
  --warmup 5 --samples 30 --case rtf_streaming_create \
  --semantic-shape large \
  --json /tmp/resprofile-after-rtf-large-0109-rN.json

taskset -c 2 heaptrack --record-only \
  -o /tmp/resprofile-heaptrack-before-xlsx-medium-0109.heaptrack \
  /tmp/litchi-perf-baseline-before-xlsx-managed-0108 --warmup 1 --samples 1 \
  --case xlsx_source_backed_cell_values_batch_edit_save \
  --xlsx-cell-crud-shape medium \
  --json /tmp/resprofile-heaptrack-before-xlsx-medium-0109.json

# The same heaptrack command was repeated for each before/after XLSX shape
# and for each before/after RTF shape, using --semantic-shape for RTF.

heaptrack_print -H /tmp/resprofile-heaptrack-<case>-0109.hist.tsv \
  /tmp/resprofile-heaptrack-<case>-0109.heaptrack.zst

taskset -c 2 perf stat -x, -e cycles,instructions,branches,branch-misses,\
cache-misses,page-faults /tmp/litchi-perf-baseline-before-xlsx-managed-0108 \
  --warmup 1 --samples 1 --case rtf_streaming_create \
  --semantic-shape medium --json /tmp/resprofile-perf-before-rtf-medium-0109.json
```

## Paired resource results

Heaptrack process totals are shown first. Peak heap and peak RSS are reported
in the exact rounded units printed by `heaptrack_print`.

| Case | Allocation calls | Allocated bytes | Temporary allocations | Peak heap | Peak RSS with heaptrack |
|---|---:|---:|---:|---:|---:|
| XLSX medium exact-256 | 8,547,384 -> 7,139,187 (-16.48%) | 1,240,794,320 -> 1,129,397,606 (-8.98%) | 1,387,508 -> 1,186,287 | 59.86M -> 59.86M | 72.01M -> 67.83M |
| XLSX dense-sparse exact-256 | 26,669,033 -> 22,930,796 (-14.02%) | 2,525,388,751 -> 2,251,515,966 (-10.84%) | 4,205,037 -> 3,675,056 | 70.60M -> 70.60M | 85.44M -> 85.30M |
| RTF medium streaming | 2,818,354 -> 655,666 (-76.74%) | 143,120,159 -> 73,914,131 (-48.36%) | 2,752,540 -> 491,557 | 24.82M -> 24.82M | 29.66M -> 33.22M |
| RTF large streaming | 45,089,098 -> 10,486,090 (-76.74%) | 2,287,130,330 -> 1,179,834,062 (-48.41%) | 44,040,208 -> 7,864,347 | 396.04M -> 396.04M | 370.03M -> 370.61M |

The five-process `/usr/bin/time -v` samples give the following medians and
ranges. RSS is maximum process RSS in KiB and is not heaptrack-instrumented.

| Case | Harness p50 median (ns) | Harness p95 median (ns) | Harness p99 median (ns) | Max RSS median (KiB; min..max) |
|---|---:|---:|---:|---:|
| XLSX medium exact-256 | 44,729,422 -> 34,155,072 | 46,682,191 -> 35,049,893 | 47,542,667 -> 35,899,442 | 65,164 (65,060..65,420) -> 65,288 (65,160..65,448) |
| XLSX dense-sparse exact-256 | 87,716,669 -> 66,646,477 | 92,393,615 -> 69,579,831 | 96,500,030 -> 69,806,544 | 79,264 (78,744..79,328) -> 79,336 (78,896..79,632) |
| RTF medium streaming | 38,111,745 -> 8,978,043 | 38,620,951 -> 9,539,556 | 38,662,248 -> 9,546,967 | 30,592 (30,592..30,720) -> 30,592 (30,592..30,720) |
| RTF large streaming | 616,709,662 -> 143,126,930 | 624,069,543 -> 145,420,909 | 628,502,057 -> 147,827,955 | 357,888 (357,760..358,016) -> 358,144 (357,888..358,144) |

The uninstrumented RSS ranges overlap. Accordingly, this record makes no
peak-RSS improvement or regression claim. Heaptrack's RTF-medium RSS increase
is profiler overhead/noise in a whole-process one-sample profile, not a
production RSS conclusion. Peak heap is flat at the displayed precision for
all four cases.

XLSX source counters stayed identical between binaries: medium has 225 source
reads, 4,230,793 source bytes, 144 ordinary-payload reads, 4,223,777 ordinary
payload bytes and six materializations per measured sample; dense-sparse has
226, 4,256,227, 145, 4,249,211 and six, respectively. Maximum in-flight reads
was one. RTF streaming retains zero output bytes and a 37-byte authoring
window. Its sink calls fall from 450,570 to 90,122 for medium and from
7,208,970 to 1,441,802 for large; the largest write is 13 -> 32 bytes.

## Identity and preservation checks

The four corpus archive hashes and output hashes were identical for every
before/after process and every measured sample:

| Corpus | Archive bytes | Archive SHA-256 | Output bytes | Output SHA-256 |
|---|---:|---|---:|---|
| XLSX medium | 4,226,429 | `dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036` | 4,226,645 | `99ef583896aa201cebcc62e205ff9979bf08b4e21df15bcd04604a2ddd14a3fd` |
| XLSX dense-sparse | 4,251,863 | `893ad3f5dd6a98aec44bc541a140048072c84c579b4b9e332431f779b097cb1a` | 4,251,968 | `484b3370ee54f2264ec8c4b4939fc0ea02e5ab3431761e9ca627dc82cfc1ca46` |
| RTF medium | 630,819 | `24e12a9c753a53a8621de2e1bd49e54da8d70143bc588316fbac80a801bd733a` | 630,819 | `24e12a9c753a53a8621de2e1bd49e54da8d70143bc588316fbac80a801bd733a` |
| RTF large | 10,092,579 | `001ee0ac2250f12cf841779fac699445567aad28d94af0c9be002ea616297435` | 10,092,579 | `001ee0ac2250f12cf841779fac699445567aad28d94af0c9be002ea616297435` |

The RTF large output hash is intentionally identical to its corpus hash; the
machine-readable summary is authoritative if this prose ever disagrees.

## `perf stat` probe

One CPU-2 RTF-medium process probe per binary succeeded with
`perf_event_paranoid=1`. Before -> after counters were cycles
911,282,809 -> 467,350,902; instructions 1,962,637,211 -> 882,214,897;
branches 361,907,150 -> 164,306,533; branch misses 1,103,019 -> 752,966;
cache misses 1,558,774 -> 1,468,670; and page faults 14,331 -> 14,334. This
was one process-wide warmup-1/sample-1 probe, so it is availability and paired
context evidence only; it is not a stable hardware-counter attribution or a
claim about any individual loop.

## Temporary artifacts and limits

The compressed heaptrack captures are temporary `/tmp` files and are not
committed. Their exact sizes and SHA-256 values are recorded in the JSON
summary. The eight captures total 2,505,794 bytes; the listed JSON, `/usr/bin/time`,
heaptrack, histogram and perf artifacts total 2,845,871 bytes. Raw reports can
be regenerated from the command templates above. No source file, benchmark
harness, aggregate performance document, or production crate was edited for
this evidence record.

The profiles do not isolate allocations by timed operation, do not establish a
peak-RSS change, and do not measure cold-cache I/O, decompressed/recompressed
bytes, concurrent scaling, or memory-bandwidth behavior. They therefore
supplement, but do not widen, the latency claims in changes 0096 and 0097.

The complete paired data, including every process-run p50/p95/p99 sample,
source/sink counters, environment, artifact hashes and command paths, is in
the [machine-readable summary](../results/xlsx-rtf-resource-profile-0109-summary.json).
