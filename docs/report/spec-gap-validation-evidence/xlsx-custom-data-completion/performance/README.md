# XLSX Custom Data completion performance

The standalone harness lives in `harness/`. It is kept outside the Cargo
workspace and resolves the library through a `source` symlink created by the
runner. This makes source selection explicit and prevents a mutable working
tree from being treated as final evidence.

Run the baseline smoke from the repository root:

```sh
bash docs/report/spec-gap-validation-evidence/xlsx-custom-data-completion/performance/run_profile.sh \
  --smoke /tmp/litchi-xlsx-custom-perf-baseline
```

The pinned baseline capture uses the detached `1d39eea51` source and a
baseline freeze receipt:

```sh
LITCHI_XLSX_CUSTOM_DATA_SOURCE_ROOT=/tmp/litchi-xlsx-custom-perf-baseline \
LITCHI_XLSX_CUSTOM_DATA_FREEZE_RECEIPT=/path/to/baseline-freeze.json \
LITCHI_XLSX_CUSTOM_DATA_CAPTURE_TOKEN=baseline-source-frozen \
LITCHI_XLSX_CUSTOM_DATA_CAPTURE_KIND=baseline \
LITCHI_XLSX_CUSTOM_DATA_RESULT_DIR=.../results/baseline-1d39eea51 \
bash docs/report/spec-gap-validation-evidence/xlsx-custom-data-completion/performance/run_profile.sh
```

The final profile requires all of the following environment variables:

```sh
LITCHI_XLSX_CUSTOM_DATA_SOURCE_ROOT=/immutable/frozen/source \
LITCHI_XLSX_CUSTOM_DATA_FREEZE_RECEIPT=/immutable/frozen/receipt.json \
LITCHI_XLSX_CUSTOM_DATA_CAPTURE_TOKEN=final-source-frozen \
LITCHI_XLSX_CUSTOM_DATA_RESULT_DIR=docs/report/spec-gap-validation-evidence/xlsx-custom-data-completion/performance/results/final \
bash docs/report/spec-gap-validation-evidence/xlsx-custom-data-completion/performance/run_profile.sh
```

The receipt must hash every selected production source file and the source
`Cargo.lock`. The runner checks those hashes before building, uses a dedicated
temporary Cargo target, records the compiler/host/binary/source/fixture
provenance, and refuses to overwrite a non-empty result directory. It records
fifteen raw samples for each of the three fixture classes and five lifecycle
lanes, with an independent `/usr/bin/time -v` sidecar for each process.

The profile compares the frozen candidate against the pinned baseline under a
matched harness and input matrix. It makes no claim about native Office files,
Office acceptance, or performance outside the authored dimensions described in
[PLAN.md](PLAN.md). Medians and percentile estimates are descriptive only;
fifteen samples do not establish tail-latency guarantees.

After both captures are present, render the matched table with:

```sh
python3 docs/report/spec-gap-validation-evidence/xlsx-custom-data-completion/performance/summarize.py \
  --baseline .../results/baseline-1d39eea51 \
  --candidate .../results/final \
  --output .../results/matched-report.md
```
