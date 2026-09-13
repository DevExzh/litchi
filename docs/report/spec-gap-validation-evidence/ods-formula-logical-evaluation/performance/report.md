# Logical evaluator performance

This batch measures the scalar logical evaluator and its lazy branches. It compares the baseline source at `3f907e9e1` with the final candidate source recorded in the candidate source manifest. The candidate adds the logical function family; the logical-function lanes are candidate-only because the baseline refuses the newly enabled functions as unsupported (`TRUE` and `FALSE` already existed). The measurements characterize these captured binaries and do not establish a general speedup claim.

## Measurement boundary

The comparable corpus has 25 cases in each of three phases (75 rows per revision). The candidate logical corpus has 45 cases in each phase (135 rows). Each process uses three warmups and 15 measured iterations. Scale lanes batch 64, 32, 8, or 2 operations for inputs of 64, 256, 1024, or 4096 terms; fixed controls batch 128 operations. The runner pins each process to CPU 2. The [environment receipt](environment.json), [runner](logical-harness/run.py), and [harness](logical-harness/src/main.rs) record the exact settings and commands.

The phases have separate meanings:

* `parse` parses a fresh expression on every operation.
* `evaluate` parses once before the timed region and evaluates repeatedly.
* `parse-evaluate` does both for every operation.

Execution setup is outside the timed region and remains live while allocator counters are read. The harness consumes each returned scalar with a full byte checksum inside the timed region, so `evaluate` and `parse-evaluate` include result consumption and should not be read as pure evaluator instruction latency. Exact expected numeric, logical, text, and formula-error values are checked before timing; a matching aggregate checksum alone would not prove value semantics.

The CSV reports p50/p95/p99 elapsed time for the whole measured batch. Values described as ns/op divide that elapsed value by the row's `repeat`. Allocation calls and requested/released bytes are likewise shown per operation when normalized; `peak_live_delta` is retained as the raw maximum for the measured batch. `max_rss_kib` is the process-level maximum resident set size.

## Reproducibility and provenance

The final paired capture ran on 2026-09-13 from 12:42:00.593–12:42:00.999 UTC for the baseline and 12:42:12.887–12:42:14.945 UTC for the candidate comparable and logical groups. The baseline executable was built from 210 source entries and has SHA-256 `412287de2566991d60f494fb0deffd768deb238c0ad5f020d8d72979ed544cc2` (1,236,720 bytes). The final candidate was built from 211 source entries, including the logical evaluator and test, and has SHA-256 `a1aaeb473345ac7e4e3c3d3e5192c86a77442d2b5dbdc00874cad6364e8265f8` (1,245,320 bytes). The [baseline source/build/binary receipts](baseline/) and [candidate source/build/binary receipts](candidate/) bind these measurements to their source snapshots. The final raw files are [baseline comparable](baseline/comparable/raw.csv), [candidate comparable](candidate/comparable/raw.csv), and [candidate logical](candidate/logical/raw.csv).

All 75 baseline rows, 75 candidate comparable rows, and 135 candidate logical rows completed with status 0. Comparable failure controls retained their expected typed failures in both revisions: unsupported reference, array, and name inputs, work/memory/object resource refusals, and cancellation.

## Comparable results

In the final single capture, the median p50 delta across the 25 rows of each phase was -0.01% for `parse`, +0.28% for `evaluate`, and +0.15% for `parse-evaluate`. The raw p50 lanes above 5% were:

| Phase and case | Baseline p50 ns/op | Candidate p50 ns/op | Raw delta |
| --- | ---: | ---: | ---: |
| `parse/control-false` | 104.1 | 110.0 | +5.63% |
| `parse/failure-reference` | 197.4 | 211.6 | +7.20% |
| `evaluate/control-coerce-256` | 25,300.8 | 27,002.6 | +6.73% |
| `evaluate/control-flat-1024` | 99,559.2 | 105,635.5 | +6.10% |
| `evaluate/control-coerce-4096` | 389,681.5 | 412,677.0 | +5.90% |

The raw p95/p99 screening flags were `parse/failure-reference` (+29.23%), `parse/failure-array` (+41.25%), `parse/failure-stack` (+24.07%), `parse/failure-cancelled` (+19.90%), `parse/control-false` (+6.43%), `evaluate/control-coerce-256` (+7.66%), `evaluate/control-flat-1024` (+5.71%), `evaluate/control-flat-4096` (+6.73%), `evaluate/control-escaped-1024` (+13.67%), `parse-evaluate/failure-memory` (+51.39%), `parse-evaluate/failure-stack` (+28.46%), and `parse-evaluate/failure-cancelled` (+74.18%). These values come from 15 measured samples per lane, and p95/p99 often select the same sample; they are retained as screening observations rather than stable tail conclusions.

The [final four-round ABAB summary](abab/summary.json) reran the flagged rows in interleaved baseline/candidate order. Only two p50 latency flags remained above 5%: `parse/control-false` at +5.22% and `evaluate/control-flat-1024` at +5.19%. The earlier `parse/failure-reference` and `evaluate/control-coerce-256` flags cleared in the repeat. `evaluate/control-coerce-4096` repeated at +3.20%, and the 1024 coercion lane was below the screening threshold before the final repeat. The final ABAB raw sequence is [here](abab/raw.csv).

The ABAB RSS results had seven clear lanes above 5%: `parse/control-false` (+6.48%), `parse/failure-cancelled` (+8.42%), `evaluate/control-utf8-left-64` (+6.18%), `evaluate/control-escaped-1024` (+7.94%), `parse-evaluate/control-utf8-left-256` (+7.64%), `parse-evaluate/control-escaped-1024` (+6.74%), and `parse-evaluate/failure-cancelled` (+8.75%). `parse/failure-array` was a 4.996% borderline case. These are whole-process peak measurements on a shared host; they do not indicate retained evaluator heap growth. The allocator counters, requested/released bytes, live-before/after values, and p50 peak-live deltas were identical for every comparable row in the final capture. The [binary size receipt](binary-sizes.json) shows an 8,600-byte file-size increase from baseline to candidate, but it does not establish a cause for the RSS variation.

The preserved [before-dispatch ABAB receipt](before-dispatch/abab/summary.json) had persistent p50 regressions of +6.46% for `evaluate/control-coerce-1024` and +6.70% for `evaluate/control-coerce-4096`. After the final conditional-dispatch frame change, the fresh capture and ABAB repeat no longer show either coercion lane above 5%. Because the captures were taken at different times, this is a measured disposition of the flagged lanes, not a causal speedup attribution. The final candidate still has the two p50 flags above and no broad latency claim is made.

The >5% checks in `docs/GOAL.md` are review triggers. This batch accepts the
remaining two latency flags and seven RSS flags as disclosed costs/uncertainties
of adding tested logical evaluation, with unchanged tracked allocation behavior.
It does not classify them as improvements or explain the RSS variation from
binary size alone. The final dispatch implementation also has fewer internal
frame variants and retains the reviewed semantics; further hot-loop changes
would require fresh evidence rather than weakening argument or resource checks.

## Candidate-only logical function measurements

These lanes establish the scale and resource behavior of the new capability; they are not a baseline-to-candidate comparison. The table gives candidate `evaluate` p50 in ns/op after parsing once and consuming the scalar result inside the timed region.

| Function | 64 terms | 256 terms | 1024 terms | 4096 terms |
| --- | ---: | ---: | ---: | ---: |
| `AND` | 3,889.5 | 11,299.1 | 43,892.6 | 198,500.5 |
| `OR` | 3,899.6 | 11,251.9 | 42,800.2 | 192,686.0 |
| `XOR` | 3,914.5 | 11,176.6 | 43,542.6 | 200,666.0 |
| `NOT` | 8,043.3 | 27,144.5 | 107,840.4 | 472,417.5 |
| `IF` | 9,357.5 | 33,305.8 | 120,534.2 | 523,472.5 |
| `IFERROR` | 7,961.9 | 27,073.2 | 106,536.8 | 451,162.0 |
| `IFNA` | 7,875.5 | 27,005.4 | 107,185.5 | 452,947.0 |

For the variadic `AND`/`OR`/`XOR` lanes, evaluate allocation calls rise from 15 per operation at 64 terms to 27 at 4096, requested bytes from 17,360 to 1,114,064 per operation, and measured batch peak-live from 8,880 to 557,232 bytes. These values show bounded, input-scaled work rather than a fixed allocation cap. The conditional-family workloads place repeated shallow calls in a flat aggregate; they measure input scaling, not arbitrarily deep recursive nesting.

Lazy and ownership cases demonstrate the semantic boundary. `lazy-if-true-heavy` remains about 427–450 ns/op across the 64–4096 scale suffixes for the unused branch, with 4 calls, 400 requested bytes, and 400 bytes of measured peak-live per operation; the heavy branch is not evaluated. The reference/error/NA lazy cases all return their expected scalar values or formula errors. `text-if-escaped` measures 576.0 ns/op with 5 calls and 403 requested bytes per operation; `text-if-concat` measures 790.5 ns/op with 6 calls and 502 requested bytes. All 135 candidate logical rows passed exact value checks and status checks. The complete candidate-only raw data is [candidate/logical/raw.csv](candidate/logical/raw.csv).

## Hardware counters

The [perf-stat receipts](perf-stat/) use `cycles,instructions,branches,branch-misses` and cover the whole process, including setup, preflight, warmups, and measured loops. They are not evaluator-only counters.

| Binary and case | Cycles | Instructions | Branches | Branch misses |
| --- | ---: | ---: | ---: | ---: |
| Baseline `control-coerce-4096` | 69,577,142 | 212,155,648 | 31,445,620 | 25,820 |
| Candidate `control-coerce-4096` | 71,223,594 | 220,962,790 | 32,964,454 | 25,581 |
| Candidate `logical-and-4096` | 38,631,886 | 124,556,517 | 20,993,952 | 39,051 |
| Candidate `logical-iferror-4096` | 82,258,800 | 277,585,718 | 46,018,483 | 38,235 |
| Candidate `lazy-if-true-heavy-4096` | 2,909,133 | 3,308,732 | 627,652 | 18,620 |

The control counter pair is consistent with a small whole-process candidate cost increase, but setup and process overhead prevent assigning those counts to one evaluator operation. The lazy counter is also consistent with skipping the unused heavy branch; it is evidence for this workload's behavior, not a general branch-cost model.

## Limits of this evidence

The comparable corpus is a regression control and the logical corpus is capability evidence; the two groups must not be combined into a single speedup number. RSS is a process-level maximum and varies with process startup and shared-host state. p95/p99 values are based on only 15 samples per lane. Timings include checksum consumption, while allocator observations are exact for the measured rows. Semantic correctness comes from the harness's exact expected-value checks and the regression suite, not from timing checksums. The final disposition is therefore bounded-memory and tested logical behavior with two small repeated p50 latency flags, variable RSS readings, and no general performance claim.
