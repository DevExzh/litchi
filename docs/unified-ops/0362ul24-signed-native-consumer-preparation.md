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
  commit `472b6705475b073e7fe08f6e11ee17fb1e29f838`.  `prepare` accepts the
  retained reviewed archive
  `/home/zhuhe/code/litchi-worktrees/.coord/w8-472b-native.V1lU2E/candidate/472b-source.tar`,
  checks every archive member against the pinned Git tree, and requires
  6,634,147,840 bytes with SHA-256
  `44c11e933974c8fa3d6a67b46739bedeef866e60fbd4360ffe7f3bb8a4dc2274`.
  If no archive is supplied, a newly generated archive is checked against this
  identity and is rejected when its tar stream differs.

Before accepting either endpoint, `source_identity` rechecks the immutable
source lock, toolchain and Cargo config.  The f2 path runs the reviewed W8
full-fence verifier over the extracted workspace (459,676 production blobs,
with only the reviewed probe, consumer adapter and pinned external lock
additions) and checks the complete source archive in one pass.  The generated
consumer's `litchi-ops` feature list is read from the fixed before/after W8
projection, and its metadata validator compares package identities,
registry checksums and dependency edges to that projection.  The two source
legs retain their role-specific locks: f2 is 139,441 bytes with SHA-256
`1b9b10ae7a9fd65db3b8b29a5bedad1aa560473c10b462f8a7d654c169335b2c`, while
the final lock is 139,606 bytes with SHA-256
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`.

`prepare --probe PATH` accepts the current external signed-native probe only
when it contains both `legacy_api` and current API branches, the XLSB
`Option`/`resolve` split, the fixture-manifest binding, and typed publication
guards.  The generated `api_compat` module is a consumer-only bridge: f2
checks the old Commit fields, `edit_with_limits`, and XLSB `Option`; final
checks the triad accessors, `edit_with_options`, and XLSB `resolve`.  No
production source tree is overlaid or modified.  `build` revalidates all of
these identities before invoking the caller's pinned offline, locked,
single-job, 20-GB-capped Cargo shim.  Endpoint capability receipts must also
bind the staged fixture manifest and pass the reviewed typed route validator
with exit code zero, no timeout or guard error, and a closed process group.

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

These checks and the preparation stage are source-only; no Cargo build, probe
execution, or native qualification is performed by this driver lane.

## Completed f2 handoff

The successful source-only f2 receipt is rooted at
`/home/zhuhe/code/litchi-worktrees/.coord/signed-build-adapter-f2-current-ppoSOO/prepared`.
The endpoint files are immutable external receipts; their exact sizes and
SHA-256 values are:

```text
adapter-api.json                 340       3aabf14f7cc620d698b7616ed508baa656c569d7c6b35ee827e40ec2d66fb85c
consumer/Cargo.toml              1,709     802163e8bae979aa36c04d89647ec4113c8532e8e199468871e388b7b0a545da
consumer/src/api_compat.rs       669       274099802d2a6c2b321a4728738d9bdee785958fdf197ac21b0d41c671b31652
consumer/src/main.rs             46,899    65e384e65a862686ec8a1eecd34a6cfdd91156367205ac9ec13ba70a47b42a00
fixture-manifest-receipt.json    445       aab530644a4c53d0607a6ca37a51effd3e8ab753d9b2b0e818b792192db78b99
plan.json                        153,884   65db7f4a5c31073c76df3d1d5856b6734433768292d78668cde5d39511d3f7df
probe-contract.json              631       3e333318400ff196d262acc44f141dda6294176caad71544d3a86a9caefcbd97
probe-input.json                 276       24f7dcc99aa6b7d8080b782aff9c13a7aff01c78bce8f02d8c3095644d053ddf
source-identity.json             150,144   80bb640b6184775ea946bc9d020eb6228ada3175a35594926b2c7e14e7301a40
```

After the coordinator grants the signed build, run the bounded build from the
integration checkout containing adapter merge
`1c2ce1f7d5230abf6c364b62a19807290ed2bc29` and protocol merge
`75e5dd4b0fc8d37e7d0c0cbc2a27e34ab5da8a50`.  Set `CARGO_SHIM` to the pinned
executable supplied by the coordinator; the driver creates fresh target and
TMPDIR directories and enforces the 20-GB cap:

```sh
set -euo pipefail
INTEGRATION=/home/zhuhe/code/litchi-worktrees/uo2-coordinator-metadata-integration
PREPARED=/home/zhuhe/code/litchi-worktrees/.coord/signed-build-adapter-f2-current-ppoSOO/prepared
CARGO_SHIM=/home/zhuhe/code/litchi-worktrees/.coord/bin/cargo
unset LITCHI_CAPPED LITCHI_BLESS
git -C "$INTEGRATION" merge-base --is-ancestor \
  1c2ce1f7d5230abf6c364b62a19807290ed2bc29 HEAD
git -C "$INTEGRATION" merge-base --is-ancestor \
  75e5dd4b0fc8d37e7d0c0cbc2a27e34ab5da8a50 HEAD
test -x "$CARGO_SHIM"
test "$(sha256sum "$CARGO_SHIM" | cut -d' ' -f1)" = \
  1b22a69fd625d57565041bf6c634d8ebd1a82be8b3d0006390472c86324230d8
env CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_NET_OFFLINE=true \
  LITCHI_MEM_MAX=20G \
  python3 "$INTEGRATION/docs/unified-ops/results/0362ul24/prepare_signed_native_consumer.py" \
  build --prepared "$PREPARED" --cargo "$CARGO_SHIM"
```
