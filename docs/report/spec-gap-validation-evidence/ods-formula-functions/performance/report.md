# OpenFormula catalog lookup and parser measurements

Baseline: `21ad9118367a7f154083ec337069896d63fb3a92` (161 names).
Candidate: the same commit plus [candidate.patch](candidate.patch), with final
formula source SHA-256 `0c3108e95e4b5fd92e616917171ebd840f26e103a2076e27592a9a330f2bd647`
and 393 standard names. The [patch replay](patch-replay.json) matches the final
root gate source manifest. Candidate checks in the isolated baseline checkout
passed 30 formula unit tests and 7 independent integration tests.

## Method and scope

The retained standalone [harness](harness/src/main.rs) uses an instrumented
System allocator, release builds, CPU 2 affinity, 3 warmup batches, and 15 measured
batches per case. Both builds use the same harness and lockfile. Each invocation
passes its input through `black_box`; long invalid lookup cases use 10,000 calls
per batch so their candidate timing is above clock resolution. Input construction,
argument processing, and reporting are outside each measured batch. Input strings
are runtime-generated, and the result checksum is observed after each batch.

Elapsed and allocation counters include token/error creation, error display, and
result destruction. Allocation requests include reallocations; released bytes
include replaced realloc storage, while deallocation call counts count explicit
deallocations. These are instrumented microbenchmarks, not uninstrumented Office
CRUD timings. No package I/O, decompression, publication, networking, evaluation,
threads, or cache refresh occurs. No end-to-end, cold-cache, scaling, or native
Office interoperability claim follows from these measurements.

Each case runs in a fresh process. RSS comes from `/usr/bin/time -v` and includes
startup, input, runtime, and static code/data, unlike the allocator's measured
peak-live delta. The [environment](environment.json), source manifests, executable
hashes, build commands, per-case commands, stdout, stderr, status, and time output
are retained. Executables and isolated worktrees are temporary and are removed
after hash verification. Baseline and candidate captures are sequential on a
shared host, not randomized trials on an exclusively reserved CPU.

With 15 batches, nearest-rank p95 and p99 both select the largest batch; they are
coarse observations, not statistically established tail bounds. Full p50/p95/p99,
mean, counters, success counts, and checksums are in [baseline/raw.csv](baseline/raw.csv)
and [candidate/raw.csv](candidate/raw.csv). Values below are normalized per call
from batch p50; they are not individually timed calls.

## Results and review disposition

All measured candidate lookup cases allocate zero heap bytes and zero allocation
calls, including lowercase/mixed case and long invalid ASCII/Unicode inputs.
The longest accepted canonical name is 19 ASCII bytes; recognition rejects an
input longer than 76 UTF-8 bytes before normalization. Unicode uppercase
compatibility remains covered by correctness tests. The lexer independently
bounds speculative ASCII name lookahead to 20 inspected identifier bytes.

Common SUM/VLOOKUP formula batches improved approximately 13–24% in p50 latency;
the bracketed absolute-reference case improved about 20%. Each removes two
allocation requests per parse. The 256-reference control changes by about +1%
and retains identical allocation requests and peak-live bytes. Short invalid
lookup changes by about +1.5%; no final comparable lane has a >5% p50 latency
regression. Malformed ASCII parser cases are 8–11% faster in the final capture.
The ordinary parser still scans and allocates diagnostics for long invalid
formulas; its cost must not be confused with bounded standalone name lookup.

The first performance review identified a malformed ASCII refusal regression.
The final code bounds speculative name lookahead and retains the previous
`to_uppercase()` diagnostic conversion. The latter has identical ASCII output
and allocation accounting and resolved the observed refusal regression on this
build. Superseded captures were removed; only the final same-harness baseline
and final candidate are retained.

RSS has one >5% review trigger: uppercase SUM is 2180 → 2308 KiB (+128 KiB,
+5.9%). Its measured peak-live delta remains 468 bytes and allocation requests
fall from 7 to 5 per call. Other processes vary by small numbers of pages and the
expanded static catalog may affect resident pages. This small process-level
increase is accepted as a measured feature cost; its cause is not isolated and
no general RSS improvement is claimed. Candidate lookup has no measured
allocator growth; this does not imply zero process memory consumption.

| Case | Calls/batch | p50 ns/call before | p50 ns/call after | Latency change | Alloc calls/call before → after | RSS KiB before → after |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| lookup-sum-upper | 10000 | 57.99 | 26.01 | -55.2% | 1 → 0 | 2192 → 2184 |
| lookup-sum-lower | 10000 | 57.90 | 49.72 | -14.1% | 1 → 0 | 2244 → 2244 |
| lookup-sum-mixed | 10000 | 57.95 | 44.81 | -22.7% | 1 → 0 | 2180 → 2264 |
| lookup-vlookup-upper | 10000 | 59.28 | 26.01 | -56.1% | 1 → 0 | 2236 → 2248 |
| lookup-vlookup-lower | 10000 | 59.38 | 50.62 | -14.8% | 1 → 0 | 2240 → 2248 |
| lookup-vlookup-mixed | 10000 | 59.36 | 46.14 | -22.3% | 1 → 0 | 2244 → 2248 |
| lookup-invalid-short | 10000 | 44.73 | 45.39 | +1.5% | 1 → 0 | 2236 → 2248 |
| lookup-invalid-ascii-4k | 10000 | 751.87 | 5.54 | -99.3% | 1 → 0 | 2244 → 2216 |
| lookup-invalid-ascii-64k | 10000 | 11145.02 | 5.54 | -100.0% | 1 → 0 | 2500 → 2244 |
| lookup-invalid-unicode-4k | 10000 | 19413.88 | 5.54 | -100.0% | 1 → 0 | 2180 → 2248 |
| lookup-invalid-unicode-64k | 10000 | 311412.26 | 5.53 | -100.0% | 1 → 0 | 2500 → 2216 |
| parse-sum-upper | 1000 | 323.73 | 246.75 | -23.8% | 7 → 5 | 2180 → 2308 |
| parse-sum-lower | 1000 | 323.54 | 269.23 | -16.8% | 7 → 5 | 2304 → 2248 |
| parse-sum-mixed | 1000 | 322.10 | 264.10 | -18.0% | 7 → 5 | 2244 → 2180 |
| parse-vlookup-upper | 1000 | 484.97 | 397.98 | -17.9% | 10 → 8 | 2188 → 2216 |
| parse-vlookup-lower | 1000 | 486.57 | 422.69 | -13.1% | 10 → 8 | 2236 → 2244 |
| parse-vlookup-mixed | 1000 | 488.05 | 422.51 | -13.4% | 10 → 8 | 2264 → 2248 |
| parse-absolute | 1000 | 374.39 | 300.74 | -19.7% | 8 → 6 | 2244 → 2248 |
| parse-cell-heavy | 128 | 18182.12 | 18363.83 | +1.0% | 265 → 265 | 2304 → 2248 |
| parse-invalid-ascii-4k | 256 | 7105.54 | 6352.10 | -10.6% | 8 → 8 | 2240 → 2232 |
| parse-invalid-ascii-64k | 4 | 127290.50 | 116543.00 | -8.4% | 8 → 8 | 2448 → 2504 |
| parse-invalid-unicode-4k | 256 | 2822.28 | 2603.57 | -7.7% | 7 → 7 | 2244 → 2248 |
| parse-invalid-unicode-64k | 4 | 40270.25 | 36137.50 | -10.3% | 7 → 7 | 2500 → 2504 |

## Additional function coverage

The six additional-name scenarios (lookup plus parsing) have no successful
baseline parser result. [candidate/new-names.csv](candidate/new-names.csv)
retains twelve absolute coverage measurements for BITAND, BIN2DEC,
BINOM.DIST.RANGE, MDETERM, UNICODE, and DDE. All calls succeed; arguments are
lexical fixtures and are not certified for arity or evaluation. These measurements
are not used to claim a before/after speedup. The integration gate independently
checks all 393 standard invocations and all 161 previous names.

## Hardware counters

A separate `perf stat` run measures the entire instrumented process for uppercase
SUM parsing (3 warmups plus 15 measured batches, 100,000 calls per batch; 1.8 million
calls total). Counters include startup, reporting, allocator instrumentation, and
teardown and cannot be attributed solely to function lookup. Both commands
succeeded with 100% counter running time; this is one process per variant, not a
statistical hardware-counter study. In particular, branch misses increase even
though instructions and cycles decrease.

| Counter | Before | After |
| --- | ---: | ---: |
| cycles | 2,626,367,280 | 1,994,176,458 |
| instructions | 6,019,072,007 | 5,034,921,894 |
| branches | 1,212,952,077 | 998,524,541 |
| branch-misses | 56,603 | 80,472 |
| cache-misses | 169,123 | 140,056 |
| page-faults | 164 | 162 |

Commands and raw counters: [perf-stat-commands.json](perf-stat-commands.json),
[baseline](baseline/perf-stat.csv), [candidate](candidate/perf-stat.csv).

## Reproduction

Create two isolated checkouts of the baseline commit. Apply `candidate.patch`
from the candidate checkout root; `git apply --unidiff-zero --check` and `git apply --unidiff-zero` both passed
in the retained replay. Copy this evidence directory's `harness/` into the same
repository-relative location in both checkouts. Build each with:

```sh
CARGO_TARGET_DIR=/your/separate/target cargo build --locked --offline --release --manifest-path docs/report/spec-gap-validation-evidence/ods-formula-functions/performance/harness/Cargo.toml
python3 docs/report/spec-gap-validation-evidence/ods-formula-functions/performance/harness/run.py --binary /your/separate/target/release/ods-formula-function-profile --output /your/results
```

Use a new output directory on each run (command logs append). The runner pins
CPU 2; adapt it if CPU 2 is unavailable and report the changed setup. Run
`harness/new_names.py` with the same `--binary` and `--output` options for the
candidate-only additional-name cases. Exact original commands and target paths
are retained with each variant. The baseline's missing function/catalog test
files are expected; the candidate adds them through the scoped patch.

From the repository root, `python3 docs/report/spec-gap-validation-evidence/ods-formula-functions/verify.py`
checks all 58 CSV rows against raw output, equal comparable inputs and checksums,
zero candidate lookup allocations, gate hashes, and baseline Git provenance
without rebuilding. The broader performance program and specification audit
remain open.
