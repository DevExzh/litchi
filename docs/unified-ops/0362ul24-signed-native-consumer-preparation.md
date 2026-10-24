# 0362ul24 signed-native consumer preparation

The signed-native probe is an external consumer.  The preparation driver in
`results/0362ul24/prepare_signed_native_consumer.py` keeps both production
source trees read-only, stages the signed fixtures into an external fixture
root, and writes only small provenance receipts in the caller's external
output directory.  It has separate `plan`, `prepare`, `build`, and `endpoint`
stages.  `endpoint` accepts only validate-child receipts bound to the actual
release binary; capability claims are not accepted as free-form metadata.

The two pinned source inputs are:

* Original f2: `/home/zhuhe/code/litchi-worktrees/.coord/w8-f2-baseline-sAirXf/src`,
  commit `f2bbbc8d07ab5da26c180f1ff74813d39f632d54`.
  Its reviewed source archive is
  `/home/zhuhe/code/litchi-worktrees/.coord/w8-f2-baseline-sAirXf/f2-source.tar`
  (6,614,528,000 bytes,
  SHA-256 `de625729ec8e6f8055e619443ad31bf1a91719cd31465975f951d54dbc1c2280`).
  The source-assembly receipt is
  `/home/zhuhe/code/litchi-worktrees/.coord/w8-f2-baseline-sAirXf/out/provenance/original-baseline-source-assembly.json`
  (SHA-256
  `b9de004ed2a74a509ebc2e3879b8d8812fcf543b92cb8bccf740ec82486ea44d`),
  and its compatibility adapter patch is bound by that receipt (153,188
  bytes, SHA-256
  `a6ffa1cc542caba1bac4d57fdf9a8e614300c8f3ac9a60716fe86dde74c7175e`).
* Final capture: `/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source`,
  commit `472b6705475b073e7fe08f6e11ee17fb1e29f838`.  `prepare` creates its
  archive with `git archive` in the external output directory and records its
  resulting artifact hash.

The driver requires the caller's pinned memory-capped cargo shim and runs
`generate-lockfile --offline`, `metadata --offline --locked`, and
`build --offline --locked --release` with `CARGO_BUILD_JOBS=1`,
`CARGO_INCREMENTAL=0`, `LITCHI_MEM_MAX=20G`, and fresh target/TMPDIR
directories capped below 20,000,000,000 bytes.  No build or native execution
was performed in this source-only lane.

Source-only checks:

```text
python3 -m py_compile docs/unified-ops/results/0362ul24/prepare_signed_native_consumer.py docs/unified-ops/results/0362ul24/test_prepare_signed_native_consumer.py
python3 -m unittest discover -s docs/unified-ops/results/0362ul24 -p 'test_prepare_signed_native_consumer.py'
```
