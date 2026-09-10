# Review request for `opc_capture_review`

Please validate the publication candidate before any files are staged.

From the repository root:

```sh
(cd docs/performance/results/full-baseline-20260910 && \
  sha256sum -c artifact-files.sha256 --strict --ignore-missing)

python3 docs/performance/results/full-baseline-20260910/scripts/derive_full_baseline_publication.py \
  verify \
  --raw docs/performance/results/full-baseline-20260910/raw/full-normal.json.gz \
  --derived docs/performance/results/full-baseline-20260910/derived/full-normal-derived.json.gz
```

The verifier must report 201 raw rows, 28 raw operation-metrics envelopes,
exactly 25 removals, three retained derived envelopes, and preserved timing /
top-level-sink projection SHA-256
`24e86e762cb85211fa7ab154a6cfa853b06041cea0e72493e4011c9eabce3371`.

For catalog validation, decompress the two normal files into a temporary
directory and run the existing corpus-binding validator. It must report 31
corpora and 201 bindings. Check the allocator report against the included V1
manifest and confirm the retained raw sidecar error is the duplicate
case/corpus binding caused by warm/cold rows sharing one corpus.

Review the README and BASELINE additions for the narrow claims: zero normal
filesystem rows, two tmpfs allocator rows, requested cache state only, and no
causal, latency-improvement, scaling, cold-cache, native-producer or
production optimization claim. Confirm the historical capture source remains
`1b3f2c2d`; publication HEAD `cf98bee3` supplies only documentation and the
verifier.
