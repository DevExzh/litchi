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
OUT=/external/.coord/w8-f2-preparation

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
projection, and an assembly manifest.  It makes no Cargo invocation and makes
no native, semantic, timing, or acceptance claim.  Generated artifacts remain
external.

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
