# Evidence: change 0574, OLE2 next-opportunity survey

Change record: [`0574-ole2-next-opportunity-survey.md`](../../0574-ole2-next-opportunity-survey.md).

Disposition: retained. `performance_claim: none`. **No production code changed in
this batch.** Everything here is attribution.

## Contents

| Path | What it is |
| --- | --- |
| `summary.json` | The folded document every table in the record is read from. Also carries the 117 per-fixture corpus rows whose raw captures are **not** retained. |
| `replay.py` | Recomputes every cited number from this directory alone. |
| `capture_counters.sh` | The driver that produced `counters/`. |
| `counters/<mode>-<op>.json` | `xls_source_attribution --warmups 20 --samples 100`, nine cells: three source implementations by three operations, on the flagship fixture. Raw per-sample elapsed and per-sample logical counters. |
| `perf/<mode>-<op>-s{100,1100}.csv` | `perf stat -x,` isolation pairs. Differencing an 1100-sample and a 100-sample child and dividing by 1000 isolates one operation, exactly as change 0564 isolated syscalls. |
| `callgrind/ann-<stem>-s{small,large}.txt` | `callgrind_annotate` self-cost output for the same isolation pairs on three fixtures: `flagship` = `ConditionalFormattingSamples.xls`, `cv` = `WithCustomViews.xls`, `54016` = `poi/.../54016.xls`. |
| `callgrind/incl-<stem>-s{small,large}.txt` | The same pairs with `--inclusive=yes`, which is where the constructor-scoped and whole-scan figures come from. |
| `callgrind/tree-flagship-s{20,220}.txt` | `--tree=caller` for the flagship pair, the source of the per-open call counts and of the `memset` caller attribution. |
| `model/globals_closure.py`, `model/globals-closure.json` | The static model of what a source-backed open's globals pass actually consumes: every four-byte header, plus the payloads of the twelve record kinds the two semantic match statements read. Deterministic; `olefile` plus the standard library. No repository code runs. |
| `model/globals_composition.py`, `model/globals-composition.json` | Per-fixture globals composition: record histogram, first and last `BoundSheet8`, SST extent. Same method. |
| `model/sst-counts.json` | `SST` declared unique and total string counts per fixture, read straight from the record payload. |
| `decision.json` | The `litchi-perf-change-decision` record: `performance_claim: none`, no code state change, reason codes, accepted costs and known gaps. |
| `ministream-probe.json` | The measured answer to "is the CFB root mini stream materialized at open", over the four fixtures whose `Workbook` stream lives in it. |

## Provenance

Commit `163ac1bd67f2a0d72c27bfec60e0a2620768cc9d`; one binary, sha256
`a2955ff9ce8ef1246ada7d8414e72fc299aaf2cbad5db2e6cd6ff85d4ab34449`, recorded in
every `counters/*.json`. A candidate implementing opportunity 1 entered the
working tree after these captures; this packet is its before leg and measures
none of it.

## Result

One source-backed open of `test-data/ole/xls/ConditionalFormattingSamples.xls`
(1,402,368 bytes) reads **53 times and 565,201 bytes** — 40.3% of the file — and
**75.9% of its wall time is spent outside the source**, not in I/O. Listing
worksheets costs nothing further: open and list are identical read for read and
byte for byte. The largest single category of that CPU is the shared-string
scan, which decodes every shared string in the workbook into a `String` and
throws it away: **82.37%** of an open of `WithCustomViews.xls` and **71.58%** of
`54016.xls`, and 34.3% of aggregate open time across 117 fixtures.

## Replay

```sh
python3 -B docs/performance/results/change-0574/replay.py
```

Rebuilding the captures needs the attribution binary:

```sh
cd tools/perf-baseline && cargo build --release --locked \
  --features xls-source-attribution --bin xls_source_attribution
docs/performance/results/change-0574/capture_counters.sh \
  tools/perf-baseline/target/release/xls_source_attribution <out-dir>
```

The models are independent of any build:

```sh
python3 -B docs/performance/results/change-0574/model/globals_closure.py --repo .
python3 -B docs/performance/results/change-0574/model/globals_composition.py --repo .
```

## What is not here

No production change, no before/after pair, no cold-cache, physical-device,
remote or range-source, peak-RSS, allocation-profile, concurrency-scaling,
real-producer or cross-platform result. No A/B/B/A latency matrix and no
host-quiescence log: the retained latency figures are single-leg attribution
medians, used to size opportunities, and are not admissible as a speedup claim.

The raw per-fixture captures behind `summary.json`'s 117 corpus rows are 4.3 MB
of elapsed samples and were discarded after folding; only the per-fixture p50 and
the logical counters survive, which is what the regression consumes.

Callgrind instruction counts for `__memset_avx2_unaligned_erms` and
`__memcpy_avx_unaligned_erms` are **upper bounds**: Valgrind instruments the
ERMS string loops per iteration, where the hardware retires far fewer
instructions. Their real cost is better read from the corpus regression's
53.4 ns/KiB coefficient than from their instruction share. Every other symbol's
share is unaffected.
