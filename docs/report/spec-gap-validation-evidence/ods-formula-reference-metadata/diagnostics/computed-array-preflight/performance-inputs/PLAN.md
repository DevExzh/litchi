# ODS reference-metadata performance profile

This profile compares the eight reference and worksheet metadata functions in
`../baseline.json` with commit `049c09cdde3978593149079c4257df047a3fa419`:
`AREAS`, `COLUMN`, `COLUMNS`, `ISREF`, `ROW`, `ROWS`, `SHEET`, and `SHEETS`.
The gate lock is the retained isolated copy with SHA256
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`; the
ambient root lock is recorded separately in `baseline.json` and is not used.

The harness keeps 28 matched controls from the preceding profiles and adds 99
candidate cases. The candidate rows cover one-area and duplicate-preserving reference lists,
ordered three-dimensional references, omitted/current coordinates, single-cell
coordinates, complete matrix ROW/COLUMN vectors, rectangular inline arrays,
runtime-kind ISREF checks, local text SHEET lookup, projected lazy IF outputs,
large descriptor geometry, direct and selected source descriptors, singleton
ReferenceList refusal for all six Reference functions, IFERROR/IFNA source
preservation and refusal, failed-first-operand fallback source policy,
array selected-source typed refusal in both cell orders, source arithmetic typed
refusal, shape refusal, a zero-read reference limit, cancellation, and one computed scalar SHEET
argument whose ABS child performs one read. The resolver contains four stable sheets in
`Main`, `Data`, `Archive`, `Hidden` order and counts every cell read; metadata
cases must report zero reads except the computed scalar SHEET lane, whose ABS
child has an exact one-read bound.

`run_profile.py` performs a fail-closed candidate preflight before timing. The
Rust harness checks the exact typed value for every scalar case and every matrix
coordinate, exact shape, typed formula errors, typed resource/cancellation
failures, and the zero cell-read bound. The matrix checks are independent of
allocator or elapsed-time measurements. The oracle and goldens under the batch
root remain separate inputs; native files are observations only.

The intended frozen capture is one coordinated run after the source handoff:

```text
python3 performance/run_profile.py \
  --candidate-root <frozen-candidate-checkout> \
  --candidate-freeze docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata/gates/freeze.json \
  --warmups 3 --samples 15
```

This produces 28 × 2 × 15 = 840 matched-control samples and 99 × 2 × 15 =
2,970 candidate samples, plus one candidate preflight and one baseline
preflight. A retry must move an existing output tree to a timestamped
`diagnostic-*` sibling and retain its receipts. Timing is still pending the
explicit source-freeze handoff; this preparation does not claim a gate pass or
performance result.
