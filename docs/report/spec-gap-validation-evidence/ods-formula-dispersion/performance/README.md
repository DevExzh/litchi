# ODS dispersion performance evidence

This directory contains the reproducible process-profile harness and receipts
for `VAR`, `VARA`, `VARP`, `VARPA`, `STDEV`, `STDEVA`, `STDEVP`, and `STDEVPA`.
The baseline is `55e147bfa0676ce6ecdc609efc682b98568b8a5f`, before the
dispersion implementation. It supplies matched controls for the existing
formula paths, including the shared statistical and database variance paths
that the candidate changes exercise.

The profile is process-isolated. A child first validates one deterministic
finite result or formula error against the typed fixture, then measures one
immutable expression repeatedly. The timed record includes elapsed time,
allocator calls and requested/released bytes, peak live bytes, execution
budget work and retained memory, resolver reads, a checksum, and external
`/usr/bin/time -v` peak RSS. Three warmups and fifteen fresh child samples are
required for every case in both evaluator phases.

The matrix covers scalar, inline-array, reference, ordered-list, 3-D,
mixed-reference/scalar, empty, formula-error, nested projected, and typed
resource-refusal inputs. `VAR`, `VARP`, and `STDEVP` list refusals are checked
as zero-read shape decisions; `STDEV` and all four A variants scan admitted
lists. VAR, VARA, and STDEVP have additional 256-row and 1024-row nested
scaling rows.

To reproduce the capture, first verify the candidate selected-file hashes with
the dispersion gate `freeze.json`, then run from the repository root during a
quiet window:

```text
python3 docs/report/spec-gap-validation-evidence/ods-formula-dispersion/performance/run_profile.py \
  --candidate-root /tmp/litchi-ods-dispersion \
  --candidate-freeze docs/report/spec-gap-validation-evidence/ods-formula-dispersion/gates/freeze.json \
  --warmups 3 --samples 15
```

The command writes `results/` with isolated baseline and candidate manifests,
raw child output, `/usr/bin/time -v` receipts, measurements, cleanup receipts,
and the profile-input hash before and after capture. Run `summarize.py` and
`verify.py` against that directory; retain those generated files with this
report. Changing any selected input after capture invalidates the receipts.
This profile makes no save, recalculation, cache-publication, native-producer,
cold-filesystem, or cross-platform timing claim.
