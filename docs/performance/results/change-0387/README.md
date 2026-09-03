# Change 0387 raw reports

This directory retains the 12 release reports used for the source-backed OPC
materialization decision and each report's schema-v2 corpus catalog.

- `control-*` and `candidate-*` use the operation-scoped counting allocator.
- `normal-control-*` and `normal-candidate-*` use the ordinary system allocator;
  their allocation status is explicitly `unavailable`.
- Each report has three warmups and 15 retained samples.
- `tiny-compressible`, `many-small-incompressible`, and
  `few-large-incompressible` are separate deterministic corpus runs.

Files are zstd-compressed without changing the underlying JSON. The
[summary](../opc-source-materialization-shared-0387-summary.json) records the
decompressed SHA-256 for every allocation-bearing report/catalog pair and the
exact source/binary/corpus bindings.

To inspect and validate one pair:

```sh
zstd -q -d -c \
  docs/performance/results/change-0387/control-tiny-compressible.json.zst \
  > /tmp/control.json
zstd -q -d -c \
  docs/performance/results/change-0387/control-tiny-compressible.corpus.json.zst \
  > /tmp/control.corpus.json
python3 tools/validate_perf_corpus_binding.py \
  --report /tmp/control.json \
  --catalog /tmp/control.corpus.json
```

The commands assume the repository root as the working directory. The
temporary paths are examples and are not part of the retained evidence.
