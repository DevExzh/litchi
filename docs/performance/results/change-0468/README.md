# 0468: remaining dense XLSX commit/save work

This descriptive profile examines the committed 0467 single-scan parser. It
changes no production or harness Rust and makes no new speedup claim. See the
[change record](../../changes/0468-xlsx-remaining-commit-profile.md),
[source review](source-review.md), and [next experiment](next-work.md).

The clean detached build is revision
`933cb6b80eaed36af21f4dce984bd6b896d3543e`. Its 6,992-file source inventory
and two compile-time fixtures match production candidate `87733cf3b`; the
intervening commit adds evidence only. The source binding records the sparse
checkout omissions. Rust 1.98.1, release debug level 1, frame pointers and
unwind tables are explicit. The build is for attribution; its uninstrumented
runs are descriptive timing context and are not a replacement for 0467 ABBA.

| Lane | Configuration | Scope |
|---|---|---|
| normal-r1, normal-r2 | 30 samples, three warmups | Uninstrumented, same diagnostic build |
| counters | 30 samples, three warmups | Whole-process perf stat |
| samples-fp | 50 samples, three warmups, 499 Hz cycles:u | Whole-process frame-pointer stacks |

Every capture uses CPU 2 and one worker. An external owned `flock` serializes
build and capture; both symbolization commands run after all four captures.
`observations.json` checks that chronology from retained command receipts.
The shared KVM host is not exclusive. `host.json` retains tool and host details.

The input is the deterministic two-sheet, 256-by-256 dense workbook with
131,072 cells and 1,311 edits. The timer covers ordinary commit and sequential
write after source opening, edit staging and sink reservation. Fixture and
expected-output construction, warmups, reopening, complete semantic readback,
reporting and teardown are included in external profiling, not in the timer.
GNU time wraps perf in instrumented lanes; their RSS includes profiler overhead.

All normal median drift is descriptive: R1/R2 p50 is 382.299942/382.735127 ms
(0.114% drift), with maximum process RSS 109,728/109,732 KiB. The profile has
15,164 samples, 132,347,593,651 weighted event periods, no reported lost samples,
and 0.592% of weight in stacks containing an unknown frame. Inclusive context
rows overlap; sampled event periods are never phase elapsed times.

`capture.py build` creates the immutable executable. To recapture, create a
fresh detached `/tmp/litchi-goal-0468/profile-tree` at the recorded revision,
with sparse paths `crates`, `tools`, `docs/adr`, `.github`, `.cargo`,
`test-data/poi/test-data/spreadsheet`, and `test-data/rtf`. Use a fresh output
bundle at the same repository depth; the helper refuses existing outputs.
Run the frozen normal-r1/counters/samples-fp/normal-r2 order under one CPU lock,
then `capture.py export` while the matching binary remains present. Gzip
replacements are authenticated before raw files are removed. Exported text
replays after binary cleanup.

```sh
python3 -B docs/performance/results/change-0468/analyze.py \
  --script docs/performance/results/change-0468/samples-fp/perf-script.stdout.gz \
  --report docs/performance/results/change-0468/samples-fp/report.json \
  --report docs/performance/results/change-0468/normal-r1/report.json \
  --report docs/performance/results/change-0468/normal-r2/report.json \
  --output docs/performance/results/change-0468/summary-fp.json \
  --additional-output docs/performance/results/change-0468/additional-summary.json
python3 -B docs/performance/results/change-0468/observe.py
python3 -B docs/performance/results/change-0468/verify.py
python3 -B docs/performance/results/change-0468/test_analyze.py
python3 -B docs/performance/results/change-0468/test_verify.py
```

Final validation passes five analyzer tests and seven verifier tamper tests,
including the live binary binding and exact derived-summary checks. The
preseal check authenticates the live binary, source tree and compile fixtures.

The analyzer imports the retained adjacent `change-0466/analyze.py`; its source
hash is included in the summary. The two source-context text snapshots are
checked against the build inventory so later source edits do not invalidate
replay. A portable copy needs this bundle, adjacent `change-0466/analyze.py`,
and `tools/perf_compare.py` at the same repository-relative paths, plus Python
standard library. The exact report-schema policy is retained as
`report-policy.json`; the initial missing-policy replay failure and correction
are retained in `validation/portable-initial-*` and `portable-correction.json`. No live Git checkout or benchmark binary is needed. `SHA256SUMS` covers every regular bundle file except itself.
The checked default matrix remains 37 cases and 201 rows, with 11 measured and
22 correctness-only representative CRUD mappings. The full non-iWork goal,
including native, cold/range, bounded-streaming and scaling evidence, remains open.
