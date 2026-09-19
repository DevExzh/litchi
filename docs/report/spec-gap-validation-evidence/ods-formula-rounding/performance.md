# ODS formula rounding profile

This directory retains a small candidate-only profile for the OpenFormula 1.4
section 6.17 rounding implementation in `litchi-ods`. It covers the direct
scalar evaluator and the matrix/value evaluator with literal arrays. The
`scalar-control` and `array-control-4x4` lanes use ordinary arithmetic as
dispatch controls. The remaining lanes exercise all eight §6.17 functions at
scalar and 4x4 array sizes. Each function has a scalar lane and a 4x4
array lane, with an additional 16x16 `ROUNDUP` scaling lane. Each phase has a
parsed-expression lane and a caller-visible parse-plus-evaluate lane.

The profile reports absolute measurements for the source snapshot recorded in
`results/source-manifest.json`. It does not contain a baseline implementation,
so it makes no before/after or speedup claim. The source snapshot includes the
rounding dispatch, scalar bridge, value evaluator, and crate manifests. A
changed source hash invalidates the retained measurements; rerun the capture
after the implementation is stabilized.

The harness performs one untimed independent numeric/shape oracle check before
each timed lane. The timed interval includes evaluator execution, checksum
folding, and result drop. Fixture generation and parsing are outside the
`evaluate` interval. `parse-evaluate` includes expression parsing. A fresh
process is used for every case and phase, with three warmups and twenty
measured samples by default. Scalar cases repeat 1,000 operations per sample;
4x4 arrays repeat 100 and the 16x16 array repeats 8. These repeats make the
short scalar operations measurable without changing the reported per-repeat
latency.

`allocator_calls_*`, requested/released bytes, and peak live bytes come from a
counting `System` global allocator in the harness. They include the evaluator
call and its per-call execution-budget setup after the counters are reset; they
are instrumentation observations rather than a claim about a production
allocator. `memory_retained_*` is the execution-budget reservation observed
while the result is live. `rss_kib` is the process maximum resident set size
from `/usr/bin/time -v`. These dimensions are reported separately because
allocator counters and budget reservations do not prove an exact RSS bound.

## Final capture

The retained capture has 19 cases (one control plus all eight functions at
scalar size, one control plus the 4x4 array functions, and a 16x16 `ROUNDUP`
stress lane), two phases, and 20 measured samples per lane. The table gives
ranges across the named lanes; `p50/op` is the p50 batch time divided by the
lane repeat count, `p95 batch` is the p95 batch time, and requested bytes and
allocator calls are p50 values divided by that repeat count. Raw per-sample
elapsed, allocation, work, and retention values remain in
`results/measurements.jsonl` and each lane's stdout receipt.

| corpus and phase | p50/op (ns) | p95 batch (ns) | allocator calls/op | requested bytes/op | retained bytes | RSS (KiB) |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| scalar, `evaluate` | 443–770 | 452,432–782,504 | 4–6 | 400–848 | 0 | 2,608–2,876 |
| scalar, `parse-evaluate` | 630–996 | 633,753–1,000,874 | 8–10 | 864–2,038 | 0 | 2,636–2,928 |
| 4x4 array, `evaluate` | 4,254–12,210 | 435,092–1,232,596 | 14–48 | 10,792–20,312 | 1,408 | 2,832–2,984 |
| 4x4 array, `parse-evaluate` | 5,016–13,278 | 507,873–1,334,316 | 28–63 | 17,260–26,823 | 1,408 | 2,820–2,968 |
| 16x16 `ROUNDUP`, `evaluate` | 169,508 | 1,369,837 | 536 | 319,832 | 22,528 | 3,060 |
| 16x16 `ROUNDUP`, `parse-evaluate` | 179,979 | 1,447,847 | 603 | 430,693 | 22,528 | 3,140 |

These are absolute observations from one release build and host. The control
lanes describe ordinary arithmetic through the same evaluator entry points;
they are included to expose dispatch and evaluator overhead and do not define
a before/after comparison.

The source receipt records the final measured production hashes:

```text
git_head: 49904ad33f35544a79c180157248c2834ab4f136
dirty_status_sha256_before: db639e0df141d18d36b544d30c562edc1e94123e5a70933e8f1dbaf43a42f810
evaluation.rs: 6eab8e35a69fa806f43c2be8fcd7a338c93ddef55837819223ac20a1da685fa3
rounding.rs: f0bdde7dad0e5d07fe2edb493be3bb8451d157b9617af87e10c34d1a1b7cdf46
value/scalar.rs: fcc077c551bba2ce5fa66207b8b952541f33c9f4e2e48f67668c4b211d806388
value.rs: 35d441d137123348f7b57a7fbb222e960c7e4bc5137eeeec25077c45bf581187
```

`results/source-manifest.json` is authoritative for the complete source
receipt, including harness hashes, and records `source_hashes_unchanged: true`.

The profile follows ADR 0005: all input is immutable and in memory, no ambient
I/O or threads are introduced, and the caller supplies finite execution
budgets. It also follows the GOAL measurement rule by retaining p50/p95/p99
sample distributions, work charges, allocation observations, peak RSS, and
the exact source/toolchain/command receipts. The harness adds no production
dependency.

Run the capture from the repository root with:

```bash
python3 docs/report/spec-gap-validation-evidence/ods-formula-rounding/run_profile.py
python3 docs/report/spec-gap-validation-evidence/ods-formula-rounding/verify.py
```

The runner builds in a fresh `/var/tmp` target directory, invokes the release
binary under `/usr/bin/time -v`, records one JSON receipt and RSS sidecar per
lane, and removes that external target after capture. `results/commands.txt`,
`results/environment.json`, `results/source-manifest.json`,
`results/measurements.jsonl`, and the per-lane sidecars are the retained
replay evidence. `results/target-cleanup.json` records that the external build
target was removed. No generated binary or build target belongs in this
directory.

Use `python3 summarize.py` to render the retained JSON lines as a compact table
when reviewing the absolute observations. The raw JSON and time sidecars remain
authoritative.
