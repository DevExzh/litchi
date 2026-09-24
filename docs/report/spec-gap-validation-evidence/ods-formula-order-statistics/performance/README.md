# ODS order-statistics performance evidence

This directory contains the reproducible process profile for `MEDIAN`, `MODE`,
`LARGE`, `SMALL`, `PERCENTILE`, `PERCENTRANK`, `QUARTILE`, and `RANK`. The
baseline is the committed dispersion implementation
`d6c0485365644a5525b7302ba0b3666d43df1b1f`.

The profile is process-isolated. Each child first validates one deterministic
fixture result or formula error, then measures one immutable expression. The
timed record includes elapsed time, allocator calls and requested/released
bytes, peak live bytes, execution work and retained memory, resolver reads, a
checksum, and external `/usr/bin/time -v` peak RSS. Three warmups and fifteen
fresh child samples are required for every case in both evaluator phases.

The bounded matrix covers scalar variadic forms, inline arrays including array
parameters, rectangular references, zero-read sequence/domain refusal,
formula-error precedence, typed resource refusal, projected 64-row reducers,
and 256/1024-row scaling. MEDIAN, MODE, LARGE, and SMALL additionally use
sorted, reverse, duplicate, and deterministic pseudo-random reference lanes.
The parameter-sensitive LARGE row uses scalar `COUNT(MUNIT([.G1:.G2]))`; its
resolver reads are kept separate from the complete reference scan.

The matrix contains 21 matched controls and 101 candidate order groups. Direct
scalar and literal-array rows are zero-read; 256-row references read 1,024
cells, 64-row references read 256, formula-error rows read 64, projected rows
read twice their complete reference, and resource or reference-list refusal
rows read zero. The MUNIT criterion row is checked separately and is expected
to read exactly 26 cells: two G-parameter reads, twelve LARGE data reads, and
twelve COUNTIF range reads.

To reproduce the final capture, verify the candidate selected-file hashes with
`gates/freeze.json`, then run during a quiet window:

```text
python3 docs/report/spec-gap-validation-evidence/ods-formula-order-statistics/performance/run_profile.py \
  --candidate-root /tmp/litchi-ods-order-statistics \
  --candidate-freeze docs/report/spec-gap-validation-evidence/ods-formula-order-statistics/gates/freeze.json \
  --warmups 3 --samples 15
```

Run `summarize.py` and `verify.py` against the resulting directory. This profile
makes no save, recalculation, native-producer, cold-filesystem, or
cross-platform timing claim.
