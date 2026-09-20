# ODS reference-metadata functions

Status: corrected source passes isolated gates, independent reviews, and
performance review; owned build/check-out cleanup is complete. Superseded snapshots are
retained separately under `diagnostics/`.

This bundle implements AREAS, COLUMN, COLUMNS, ISREF, ROW, ROWS, SHEET and
SHEETS in the existing evaluator. Direct reference metadata requires zero cell
reads; matrix axis output and computed arguments retain bounded evaluation.
External value access, named expressions, workbook recalculation, and formula
cache publication remain outside this batch.

See [completion.md](completion.md) for behavior, validation, measured limitations,
and cleanup; [contract.md](contract.md) defines the normative implementation
profile. Independent [semantic](semantic-review.md) and
[resource/cache](resource-review.md) reviews are bound to the frozen sources by
[review-receipt.json](review-receipt.json).

The corrected isolated gates pass seven checks and 1,655 tests. The independent
oracle contains 91 observations; the native fixture compares 32 rows with four
documented host divergences. The superseded capture retains 4,650
samples and the final capture 4,920. Matched allocation/work/read/output
accounting is unchanged, with explicitly accepted evaluate-only SUMIFS timing
regressions of +5.865% and +6.284%, respectively. This is
not an end-to-end speedup claim.

Run retained-evidence verification from the repository root:

```sh
python3 -B docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata/verify.py
```

The verifier checks frozen source and Git custody, reviews, oracle/native
provenance, gate commands and logs, and raw performance samples. It uses the
retained isolated Cargo.lock; the distinct ambient lock remains untouched.
Gate staging/reproduction is documented in [gates/README.md](gates/README.md),
and profile reproduction in [performance/README.md](performance/README.md).
The earlier source-policy gate snapshot under `diagnostics/` is retained only
as diagnostic evidence, not as acceptance of the final implementation.

For a new capture, use the full command from the repository root (the freeze
path is resolved relative to the command's working directory):

```sh
python3 docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata/performance/run_profile.py \
  --candidate-root /absolute/path/to/frozen/checkout \
  --candidate-freeze docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata/gates/freeze.json \
  --warmups 3 --samples 15
```

Preserve existing capture output before any explicitly justified new run.
