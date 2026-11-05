# 0362hf — pristine f2 W8 before-leg preparation

This record prepares, but does not qualify, the original W8 before endpoint.
The production source is the unchanged full tree at
`f2bbbc8d07ab5da26c180f1ff74813d39f632d54`.  The source archive and external
workspace are checked against every f2 blob and path.  No later production
commit, API patch, or workload replacement is accepted.

The corrected 109-route probe is pinned to its independently reviewed blob at
`d39d179353b5fd1ce82ca479b16cee2fc4dc8ffb`:

```text
path: crates/litchi-ops/examples/file_w8_qualified_probe.rs
bytes: 74747
sha256: 3a01dd2cbb971b4eba54e2eccffa99e1b3c4557db14af03e61869f99f0a0cc27
```

The f2 consumer API exposes `Commit.file` and `BatchCommit.file`.  The
standalone adapter at
`crates/litchi-ops/examples/file_w8_original_baseline_probe.rs` is generated
from the corrected probe and changes exactly four consumer expressions:

```text
unified_batch   BatchCommit.file
unified_import  Commit<Summary>.file
unified_merge   Commit<Summary>.file
unified_refresh Commit<Summary>.file
```

Every other byte of the probe is retained.  The adapter is an added example,
not a production source change.  The workspace checker permits only that
adapter, the unchanged corrected probe, and `Cargo.lock` as additions.  f2 has
no tracked root lockfile, so the lock is an explicit, separately pinned
external resolution whose metadata records f2 as its source commit and a
non-empty origin description.  It is the only intentional dependency input
delta.

## Preparation

All source and generated paths below are outside the checkout.  First create a
pristine f2 archive and extract it without modifying its production files;
then copy the reviewed corrected probe into its fixed example path and supply
the separately pinned lockfile:

```sh
set -eu
ROOT=/home/zhuhe/code/litchi-worktrees/uo2-w8-readiness
SRC=/external/w8-f2-workspace
ARCHIVE=/external/w8-f2.tar
LOCK="$SRC/Cargo.lock"
OUT=(deleted-coord-artifact)

git -C "$ROOT" archive --format=tar --output="$ARCHIVE" \
  f2bbbc8d07ab5da26c180f1ff74813d39f632d54
mkdir -p "$SRC"
tar -xf "$ARCHIVE" -C "$SRC"
git -C "$ROOT" show \
  d39d179353b5fd1ce82ca479b16cee2fc4dc8ffb:crates/litchi-ops/examples/file_w8_qualified_probe.rs \
  > "$SRC/crates/litchi-ops/examples/file_w8_qualified_probe.rs"

python3 "$ROOT/tools/w8_original_baseline.py" prepare \
  --git-root "$ROOT" \
  --workspace "$SRC" \
  --probe "$SRC/crates/litchi-ops/examples/file_w8_qualified_probe.rs" \
  --lock "$LOCK" \
  --lock-source-commit f2bbbc8d07ab5da26c180f1ff74813d39f632d54 \
  --lock-origin "pinned offline resolution generated from the pristine f2 workspace" \
  --source-archive "$ARCHIVE" \
  --output-dir "$OUT" \
  --target x86_64-unknown-linux-gnu \
  --profile release
```

The command emits a canonical fixture descriptor, a 109-row route map with
route-specific native/facade timing and publication bindings, a dependency
projection, a source-assembly record whose compatibility patch contains only
the two added probe/adapter files, and an assembly manifest.  It makes no
Cargo invocation and makes no native, semantic, timing, or acceptance claim.
Generated artifacts remain external.

After coordinator review, the only compile candidate is the standalone
adapter in the external f2 workspace, with one build job and the exact W8
all-features selection (the probe's 11 format features are enabled by that
selection):

```sh
CARGO_BUILD_JOBS=1 cargo check --offline --locked -p litchi-ops \
  --example file_w8_original_baseline_probe \
  --all-features
```

Untimed 109-route validation and quiet-host ABBA timing are separate pending
steps.  Their receipts must bind the same f2 tree, adapter, corrected probe,
fixture, lock, feature selection and toolchain before any qualification status
is considered.

## Compile and untimed validation receipt

The one permitted original-f2 build and the untimed validation have now been
run in the external workspace
`(deleted-coord-artifact)`.  This is a
before-leg capability/fidelity receipt only; it makes no native qualification,
acceptance, or performance claim.

The release build used one job, `--offline --locked --all-features`, debug
information disabled, the 20 GB cargo memory cap, and a workspace-local
temporary directory.  It succeeded in 22m47s.  The executable is retained
outside Git at
`target/release/examples/file_w8_original_baseline_probe` (65 MiB,
SHA-256 `f4c5a8ed31888f04337466265e03a8d02fc2df14fbac0c3eb3d6be83c358676d`).
The compact build receipt is
`out/provenance/original-baseline-build-receipt.json` (3,993 bytes,
SHA-256 `621c5a75a4ef57d7ffcc4f1b4168a23ce9b76b13e88acb4b7d3c592ae3f2abd7`).

The probe then ran each of the 109 routes once with `--mode validate`, without
calibration or timing mode.  Every route reported semantic equality,
semantic-state validity, exact source preservation, paired leaf/facade
read/output fidelity, nonempty payload hashes, and repeated publication
equality.  The receipt is
`out/provenance/original-baseline-untimed-validation.json` (71,489 bytes,
SHA-256 `4277f5f4d3518bc58cd6ab60e56c45f4a3e3ede95752427fa3f12ffb2df6d1c`).
The existing fail-closed endpoint verifier accepted the complete source,
lock, fixture, build, and validation identity; its compact endpoint manifest
is `out/provenance/original-baseline-endpoint.json` (3,192 bytes,
SHA-256 `56543e806b774716c9e381b3da33400dc99ff337a82031bc0eb8409951dbc608`).

The source archive remains external (6,614,528,000 bytes,
SHA-256 `de625729ec8e6f8055e619443ad31bf1a91719cd31465975f951d54dbc1c2280`).
The deterministic fixture and locked registry-resolution identities are
recorded separately in
`out/provenance/original-baseline-seed-registry-identity.json` (1,883 bytes,
SHA-256 `26354ad50b20bd59969178da2462f3155270e72c17619c5ecd3b8ab1bcb41838`).
Raw per-route JSON, the source archive, target directory, and executable stay
outside the repository.  No candidate build or ABBA timing was run.
