# 0960 strict-lane RawCell reservation screen

The complete second epoch and rejected screening decision are recorded in
[epoch2](epoch2/README.md). The following describes the historical first epoch.

This is a bounded provenance and quality packet for the quality-frozen
candidate and root-authorized screening epoch. The capture epoch is incomplete:
root session `57382` has no lane-terminal receipt. The retained native packet
has 34 child directories/start records, 33 launch/terminal/validation/post-
terminal records, and one child without a launch or terminal receipt
(`native__medium-one-edit__b02__p1`) with only `start.json` and an empty
`time-v.txt`; the authorized schedule expected 36 children. The allocator lane was never started. The complete
per-file inventory is `review/partial-capture-interruption-v1.json` (SHA-256
`f68a039593fb1ce6715f06bb5246d2c1e438c65384577f171236036e85605ca1`). No
result, reduction, timing conclusion, or adoption claim is included. Candidate
source-custody v3, the final plan, root GO, build/retention receipts, and
preflight controls are bound in `protocol/`; raw captures, logs, binaries, and
full source manifests remain external.

The quality evidence currently retained is:

- the full `litchi-xlsx` library test: `rc=0`, 1,404 passed, from the
  pre-final-quality codec hash `dd85fd4fb42dfc770d50510613c1d0585bcab63384ffb7ec1b5a8ac60643299e`;
- focused strict-lane tests v7: `rc=0`, 7 passed and 1 ignored, from corrected
  codec hash `19d8be4a882fe1ed8c80d428fdf5780e30b2f979bb22841cf5eef540e3153657`;
- Clippy v2 for `litchi-xlsx --lib -D warnings`: `rc=0`, from the same
  corrected codec hash;
- scoped rustfmt v1 for the changed codec: `rc=0`, from the corrected codec
  hash; and
- the managed route census: `rc=0`, eight worksheet inputs covering 27,008
  cells. Its receipt predates the corrected codec hash and is retained with
  that exact source identity.

The full test and route-census receipts are therefore historical quality
inputs, not a claim that the complete suite was rerun after the final
needless-borrow correction. Final-source-review-v2 records this boundary; the
post-fix focused, Clippy, and scoped-format checks are bound to codec hash
`19d8be4a882fe1ed8c80d428fdf5780e30b2f979bb22841cf5eef540e3153657`.

The external artifact manifest is
`review/external-artifact-manifest-v1.json` (SHA-256
`e57b47020b0a852288afc3b15e975ddd9d08b7b51745bbece4a12b611663a548`). It records sizes, modes, SHA-256
values, regeneration commands, historical failure logs, both pre-quality v1
source manifests, the quality-frozen v2 source manifests, the retained 0957
source and binaries, the candidate patch, and the eight worksheet route census
directory. Large source manifests, source trees, logs, route inputs, and
executables stay outside this repository. The v2 directory census matches the
quality-frozen source manifests exactly; v1 remains listed as historical input
provenance.

The copied `quality/` files are bounded command, environment, terminal, and
manifest checkpoints. The copied `protocol/` files bind final plan SHA
`a9c0db61ae9dd48c349d3c264c140d6126c35979c32cecc1f671ded8bb2b176d`, root GO,
runner v7, reducer v3, selector/rule v2, lifecycle controls, and actual
preflight. The candidate release executables are described by SHA in the
external manifest and are never copied here. Full logs and source trees are external. The copied
reviews include the design and screening controls; `review/final-source-review-
v2.md` is the current source review and `review/historical-final-source-review-
v1.md` is retained as superseded pre-correction context. Superseded
Clippy/focused failures and the blocked pre-freeze review remain listed by
descriptor in the external artifact manifest rather than being copied as
redundant logs.

To regenerate the external descriptor in a fresh destination (never overwrite
an existing output):

```sh
python3 -B docs/performance/results/change-0960/tools/build_external_artifact_manifest_0960_v1.py \
  --output /tmp/manifest-0960-v1.json
```

The source snapshot command recorded by the custody receipt is:

```sh
git -C /home/zhuhe/code/litchi archive --format=tar \
  a6223afa0899d5800d21ed9dc6c49304fbb2e8ca -- . \
  ':(exclude)docs/performance/**' ':(exclude)docs/report/**'
```

For the quality-frozen v2 BOM, use the fresh extracted roots and the retained
lock overlay, apply the bounded candidate patch, then regenerate each BOM with
the copied helper (the output path must not already exist):

```sh
python3 -B docs/performance/results/change-0960/tools/rebuild_source_manifest_0960_v1.py \
  --root /home/zhuhe/litchi-performance-0960/source-baseline-v1 \
  --output /tmp/source-manifest-baseline-v2.json --role baseline \
  --source-commit a6223afa0899d5800d21ed9dc6c49304fbb2e8ca
python3 -B docs/performance/results/change-0960/tools/rebuild_source_manifest_0960_v1.py \
  --root /home/zhuhe/litchi-performance-0960/source-candidate-v1 \
  --output /tmp/source-manifest-candidate-v2.json --role candidate \
  --source-commit a6223afa0899d5800d21ed9dc6c49304fbb2e8ca
```

The benchmark result summary and reduction are intentionally absent here:
the capture did not reach a lane-terminal state and the allocator lane did not
start. The authorized final plan is present in
`protocol/plan-final-v1.json`; `protocol/root-live-state-v1.json` is retained as
historical pre-interruption authorization state, not terminal evidence. The
partial packet inventory is the only capture-state evidence in this repository;
no timing fields from retained child reports were analyzed. The independent
read-only continuation audit is `review/continuation-independent-audit-v1.json`,
and `review/continuation-write-limits-v1.json` records that repository staging
was unavailable because `.git` is read-only in that continuation environment; no
commit or external cleanup was performed. The candidate patch is present as
source provenance.
