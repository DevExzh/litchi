# ODS order-statistics performance plan

This profile measures the eight bounded OpenFormula ordering reducers
`MEDIAN`, `MODE`, `LARGE`, `SMALL`, `PERCENTILE`, `PERCENTRANK`, `QUARTILE`,
and `RANK` against the frozen baseline
`d6c0485365644a5525b7302ba0b3666d43df1b1f`.

The baseline and candidate use the retained isolated dispersion gate lockfile.
The matched controls include arithmetic, `SIN`, `IMSUM`, `SUM`, `SUMIFS`,
`AVERAGE`, `COUNTA`, `VAR`, `STDEV`, `DVAR`, and `DSTDEV`, plus the existing
array and reference controls. Candidate-only rows cover scalar variadic calls,
inline arrays and lifted array parameters, sorted/reverse/duplicate/random
reference lanes for the selection functions, typed/domain and sequence-shape
refusal, formula errors, resource refusal, projected reducers, and 256/1024
row scaling. A dedicated projected `MUNIT` parameter row makes complete
reference caching distinguishable from position-sensitive parameter lifting.

Every group uses three untimed warmups and fifteen fresh child processes in
both `evaluate` and `parse-evaluate` phases. Each child performs one untimed
correctness preflight, then records elapsed time, peak RSS, allocator calls and
bytes, live-byte balance, evaluator work, retained budget memory, resolver
reads, and a result checksum. Capture is valid only after the candidate freeze
matches all selected files, including the numeric-oracle generator and order
contract, and during a quiet window.

The final case matrix has 21 matched controls and 101 candidate order groups
(122 candidate groups total). The eight reducers each have scalar, inline-array,
type/domain refusal, sorted 256-row reference, formula-error, projected 64-row,
and resource-refusal rows. LARGE and SMALL add inline array-rank rows. MEDIAN,
MODE, LARGE, and SMALL add sorted-64, reverse-64/256, duplicate-64/256, and
pseudo-random-64/256 rows. MEDIAN, LARGE, SMALL, PERCENTILE, PERCENTRANK,
QUARTILE, and RANK add projected 256/1024-row rows. The final extra row is a
projected LARGE criterion using `COUNTIF(...;LARGE(...;COUNT(MUNIT([.G1:.G2]))))`.

Direct scalar and literal-array rows read zero resolver cells. Sorted 256-row
references read 1,024 cells; the 64-row pattern rows read 256; their 256-row
pattern counterparts read 1,024; error rows read 64 while retaining the error;
projected rows read twice their complete reference (`512`, `2,048`, or `8,192`);
and resource-refusal rows read zero. MODE and QUARTILE reference-list refusal
rows are also zero-read. The projected MUNIT criterion row is expected to read
exactly 26 cells: two projected G-parameter reads, twelve LARGE data reads,
and twelve COUNTIF range reads. It is kept separately from the complete
reference bounds.

```text
python3 docs/report/spec-gap-validation-evidence/ods-formula-order-statistics/performance/run_profile.py \
  --candidate-root /tmp/litchi-ods-order-statistics \
  --candidate-freeze docs/report/spec-gap-validation-evidence/ods-formula-order-statistics/gates/freeze.json \
  --warmups 3 --samples 15
```

Run `summarize.py` and `verify.py` after capture. Retain raw child output,
`/usr/bin/time -v` receipts, manifests, and cleanup receipts. Any selected
input change after capture invalidates the receipts.
