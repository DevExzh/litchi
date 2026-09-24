# Descriptive-statistics performance evidence

This directory defines the reproducible profile for `AVEDEV`, `DEVSQ`,
`GEOMEAN`, `HARMEAN`, `KURT`, `SKEW`, and `SKEWP`. The comparison baseline is
`b8e5d5fe257fd95747c69a3c44a53cedd96f77ed`, using the retained
order-statistics gate lock at
`docs/report/spec-gap-validation-evidence/ods-formula-order-statistics/gates/Cargo.lock`.

The final profile will compare 24 matched controls, including the representative
scalar controls (`MEDIAN`, `RANK`, and `PERCENTRANK`), with the
bounded descriptive matrix in [PLAN.md](PLAN.md). It will use three warmups and
fifteen fresh process samples in both evaluator phases. Raw stdout, external
`/usr/bin/time -v` receipts, allocator/work/read fields, source closures, and
cleanup receipts are retained beside the summarized report.

The reviewed contract input is SHA-256
`6d0127bbf1867553860ba20013aab530c9931ea961e022532e9c16df4424965f`. The
harness lists and runs untimed correctness preflights for every selected case;
the runner checks the contract hash, builds with the retained lock, and records
those preflight receipts before any timing loop. Timing evidence is written
under `results/` by the authorized command; these protocol files remain
unchanged. Root must provide the frozen candidate manifest and quiet-window
authorization before the command below is run; after that point the profile
inputs and source closure are immutable.

The authorized capture command is:

```text
python3 docs/report/spec-gap-validation-evidence/ods-formula-descriptive-statistics/performance/run_profile.py \
  --candidate-root /tmp/litchi-ods-descriptive-statistics \
  --candidate-freeze docs/report/spec-gap-validation-evidence/ods-formula-descriptive-statistics/gates/freeze.json \
  --warmups 3 --samples 15
```

Run `summarize.py` and `verify.py` only after capture. The profile makes no
save, recalculation, native-producer, cold-filesystem, cross-platform timing,
or unsupported numerical-semantics claim.
