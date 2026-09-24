# Historical full baseline publication: 2026-09-10

This directory publishes one descriptive historical control capture. The
capture source is commit `1b3f2c2d0c059e8e59272775a97d0567e14e67f2`, while the
portable publication candidate was prepared from feature HEAD
`cf98bee37455e8e6d0e73002ff2e94299ef546cb`. The feature HEAD identifies the
publication documentation and verifier; it was not the measured source.

The normal report has 201 rows across 37 cases and 31 deterministic corpora,
with 15 samples and three warmups per row. It contains zero normal filesystem
rows. Its raw JSON is retained as `raw/full-normal.json.gz`; the uncompressed
JSON SHA-256 is
`25800c3b44e175556ad5aacb3172d9be3eb021521ac4b1720c31c59b22f04b10` and its
case/corpus identity digest is
`f0fd76293959e72211e06e51b0a2b41f371423fda34c55144077193d262b1670`.

The capture emitted 28 normal `operation_metrics` envelopes. Twenty-five
save-row envelopes have `sample_indices` in ordinal order while elapsed
samples are reported in a different order; their constant sink vectors do not
support per-sample operation correlation. The raw report remains untouched.
`derived/full-normal-derived.json.gz` is a reproducible derived view that
removes only those 25 envelopes. It preserves every row's elapsed timing and
top-level `sink`, all report-level fields, and the three aligned operation
envelopes. Its uncompressed JSON SHA-256 is
`35d90ffd8376d371e34a6aa5cfde6e64ff24413736e53bb8a5dc65b9af5c67d0`.

Reproduce and verify that transformation from the repository root:

```sh
python3 docs/performance/results/full-baseline-20260910/scripts/derive_full_baseline_publication.py \
  derive \
  --raw docs/performance/results/full-baseline-20260910/raw/full-normal.json.gz \
  --derived /tmp/full-normal-derived.json.gz \
  --receipt /tmp/full-normal-derived-receipt.json

python3 docs/performance/results/full-baseline-20260910/scripts/derive_full_baseline_publication.py \
  verify \
  --raw docs/performance/results/full-baseline-20260910/raw/full-normal.json.gz \
  --derived /tmp/full-normal-derived.json.gz
```

The checked allocator artifact is a separate two-row smoke over
`opc_file_eager_open` and the `few-large-incompressible` corpus, with
`warm`/`cold-requested` rows and `system_allocator_operation_scoped` metrics.
The report recorded a `tmpfs` filesystem. It is allocation evidence only, not
full allocation coverage; 198 allocator regions remain outside this artifact.
The allocator report's uncompressed SHA-256 is
`cbc36473c73c631c74c15dc3f4ebd2db19c6278af42ae0f16634e04b56c22c7c` and its
two-row identity digest is
`debb700930d4be1d2d597de60ef5b64200403407c138bab17dd268996975ec7a`.

The optional allocator V2 corpus sidecar failed deterministically because its
catalog writer treated the warm and cold rows sharing one corpus as a duplicate
case/corpus binding. The failed sidecar log is retained under `raw/`. The
allocator report was run once more without that optional sidecar and checked
against `raw/opc-file-allocator.manifest-v1.json`; no raw timing report was
rewritten or recaptured.

The original capture plan, including its historical wrapper name, remains
under `raw/capture-plan-historical.md`. The successful post-review command is
the corrected `raw/run-full-baseline-and-allocator-v2.sh`. The reviewed lock
patch and generated lock are included under `locks/`. Source identity,
toolchain, build flags, binary identities, run handles, failed sidecar,
manual allocator rerun, and contention snapshots are documented in
`provenance/historical-capture-provenance.md`; the raw-to-derived receipt is
`derived/full-normal-derived-receipt.json` and the complete stored-file hash
receipt is `artifact-files.sha256`. The separate publication and historical
source identities are in `provenance/publication-source-identity.txt`.

`warm` and `cold-requested` are requested harness states; no physical cache
eviction claim is made. The host was shared and contention snapshots are
retained. These artifacts make no causal claim, latency improvement claim,
scaling claim, native-producer claim, or production optimization claim.
