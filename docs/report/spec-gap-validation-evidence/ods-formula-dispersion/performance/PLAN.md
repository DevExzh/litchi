# ODS dispersion performance plan

This profile measures the eight bounded OpenFormula dispersion reducers
`VAR`, `VARA`, `VARP`, `VARPA`, `STDEV`, `STDEVA`, `STDEVP`, and `STDEVPA`
against the frozen baseline `55e147bfa0676ce6ecdc609efc682b98568b8a5f`.
The baseline and candidate use the retained isolated gate lockfile from the
core statistical-reducer batch. The candidate source closure and authored
profile inputs must be frozen before capture.

The matched matrix contains the existing arithmetic, `SIN`, `IMSUM`, `DSUM`,
`SUM`, `SUMIFS`, `AVERAGE`, `COUNTA`, `DVAR`, and `DSTDEV` paths. The candidate
matrix adds ten rows per dispersion function: scalar, inline array, rectangular
reference, ordered reference list, 3-D reference, mixed reference/scalar
arguments, empty selection, formula-error selection, a 64-row projected
reducer, and a typed resource-refusal row. `VAR`, `VARA`, and `STDEVP` add
256-row and 1024-row projected scaling rows, producing 86 candidate-only
workloads and 103 total candidate workloads including the 17 controls.

Reference rows use a deterministic borrowing resolver. Read counts are
checked against the ordered traversal contract: list shape refusals for
`VAR`, `VARP`, and `STDEVP` read zero cells; `STDEV` and the A variants scan
their admitted lists; 3-D rows read one cell per sheet; nested rows read the
outer condition and cached inner reference exactly once per output cell; and
typed resource rows read zero cells.

Each final group uses three untimed warmups and fifteen fresh measured child
processes in both `evaluate` and `parse-evaluate` phases. Every child performs
one untimed correctness preflight, then records elapsed time, peak RSS,
allocator calls and bytes, live-byte balance, evaluator work, retained budget
memory, resolver reads, and a result checksum. Run the capture only after the
candidate selected-file hashes match `gates/freeze.json` and during a quiet
window:

```text
python3 docs/report/spec-gap-validation-evidence/ods-formula-dispersion/performance/run_profile.py \
  --candidate-root /tmp/litchi-ods-dispersion \
  --candidate-freeze docs/report/spec-gap-validation-evidence/ods-formula-dispersion/gates/freeze.json \
  --warmups 3 --samples 15
```

Then run `summarize.py` to produce the human-readable and JSON summaries and
`verify.py` with the same candidate root and freeze file. Keep the generated
`results/` manifests, raw child output, timing receipts, and cleanup receipts
with the report; any selected-input change after capture requires a new run.
