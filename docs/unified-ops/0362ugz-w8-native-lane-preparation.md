# 0362ugz — frozen W8 File native lane preparation

This is an input/provenance preparation receipt for the fixed W8 source
freeze `f6fe430a4d19d94aa360c0e9ed03ebd757265225`.  It does not contain a Cargo
build, a binary, an untimed native run, timing, ABBA data, or an acceptance
claim.  The historical f2 endpoint at
`(deleted-coord-artifact)` was read only; none of its endpoint evidence was
relabelled.

The receipt is deliberately pinned to f6fe.  An earlier b673 candidate lane
was prepared before this final source freeze; its candidate artifacts are not
relabelled.  The f2 artifacts below are reused only because their f2 source,
probe, lock and fixture inputs are byte-identical; the f6fe candidate source
archive and source assembly were generated afresh.  If the final freeze moves,
regenerate the fixture, projection, source archive and both endpoint inputs as
a new lane.

## Frozen inputs and external lane

All generated files below are outside Git in
`(deleted-coord-artifact)` (the lane).
The lane has separate `target-f2`, `target-f6fe`, `tmp/f2-build`,
`tmp/f6fe-build`, `validation/{f2,f6fe}`, `endpoint/{f2,f6fe}` and `abba`
directories reserved for later coordinator-controlled work.  No target
contains a build from this preparation.

| input | bytes | SHA-256 |
|---|---:|---|
| candidate `crates/litchi-ops/examples/file_w8_qualified_probe.rs` | 75,035 | `21c8e3833c96411a612a2631f98093bd70cf0a5c28b2952c3d15873f464e54b0` |
| shared candidate fixture descriptor | 4,231 | `b7439d899a940b4135dd44081554a4aa9d7e70c19e2ccd3339440baad5ca7e09` |
| shared candidate dependency projection | 146,453 | `586cbbe06a8591322488a9650bbc098ad4b9f2f137c95e2ea80617728cbcd2e5` |
| copied W8 lock (`Cargo.lock`) | 139,591 | `52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6` |
| candidate source archive (`f6fe-source.tar`) | 6,629,089,280 | `cf533108be38b84f8c2ac4de536d6196e16c701c84f9a70645d05810db6d7f3b` |
| candidate source assembly | 472 | `0da4cee1a1fd857ec856dad9b8499e534a5bcef3fd2b44a827c9c6e8cf24bfc5` |

The candidate toolchain file is 86 bytes, SHA-256
`e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310`.
The candidate workspace `Cargo.toml` is 10,672 bytes, SHA-256
`f5ec007892023b38f1deaef1ce20714db58b58448a3bea088be3b847b84570de`, and
the `litchi-ops` manifest is 13,489 bytes, SHA-256
`b22e9ff3a31366cd0bf85bd0d65b8add84cb89c612535ae905ea578b9a0eaf1d`.
The projection selects all features for `litchi-ops`, target
`x86_64-unknown-linux-gnu`, release profile, and the complete 109-route
workload.

The future candidate endpoint is the separate clean detached worktree
`/home/zhuhe/code/litchi-worktrees/uo2-w8-f6fe-native-endpoint`, at the same
f6fe commit.  Its external `Cargo.lock` is 139,591 bytes with SHA-256
`52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`.
The earlier provenance-generation receipt used the gate checkout only to
generate the already-recorded fixture/projection hashes; no future candidate
build, validation or endpoint command uses that checkout.

The build-free provenance test was run against the frozen candidate:

```text
python3 -m unittest tools.test_w8_file_provenance
Ran 12 tests in 0.078s
OK
```

Its f6fe-bound external log is 111 bytes, SHA-256
`ca18cc8e8cf94f56e58232f50620b1e8cacd8cccd6a6b7690d4968cc5d62d3ac`.

The f6fe fixture/dependency command stdout artifacts are also retained in the
lane: `commands/fixture.stdout` is 214 bytes with SHA-256
`583826536fd40f3066620e7fb76037bb32bc9008a54495fc129562bee714f24a`, and
`commands/dependency.stdout` is 219 bytes with SHA-256
`bb40d6a0c62b0dc47bca35dd18f8e92d1b2619f460c1b815b8ca59cdc9196f61`.
The source-assembly stdout is 205 bytes with SHA-256
`d531c35b71d09f263b6f8bc90d715a0d49983cbc67333eb82a79b606e2081fb3`.
These commands were run against the f6fe tree and did not invoke Cargo:

```sh
C=/home/zhuhe/code/litchi-worktrees/uo2-budget-validation
LANE=(deleted-coord-artifact)
PROBE="$C/crates/litchi-ops/examples/file_w8_qualified_probe.rs"
python3 "$C/tools/w8_file_provenance.py" fixture \
  --probe "$PROBE" --output "$LANE/provenance/litchi_w8_fixture_descriptor.json"
python3 "$C/tools/w8_file_provenance.py" dependency \
  --workspace "$C" --package-manifest "$C/crates/litchi-ops/Cargo.toml" \
  --lock "$C/Cargo.lock" --toolchain "$C/rust-toolchain.toml" \
  --probe "$PROBE" --fixture "$LANE/provenance/litchi_w8_fixture_descriptor.json" \
  --all-features --target x86_64-unknown-linux-gnu --profile release \
  --output "$LANE/provenance/litchi_w8_dependency_projection.json"
git -C "$C" archive --format=tar \
  --output="$LANE/f6fe-source.tar" f6fe430a4d19d94aa360c0e9ed03ebd757265225
python3 "$C/tools/w8_file_provenance.py" source \
  --source "$LANE/f6fe-source.tar" \
  --source-commit f6fe430a4d19d94aa360c0e9ed03ebd757265225 \
  --overlay-kind none \
  --description "fixed W8 source freeze f6fe; no production compatibility overlay" \
  --output "$LANE/provenance/f6fe-source-assembly.json"
```

## Fresh f2 before-leg assembly

The historical source archive is
`(deleted-coord-artifact)/f2-source.tar`,
6,614,528,000 bytes, SHA-256
`de625729ec8e6f8055e619443ad31bf1a91719cd31465975f951d54dbc1c2280`.
Its source commit is
`f2bbbc8d07ab5da26c180f1ff74813d39f632d54`, with 459,676 production files
and production-tree SHA-256
`23e332254abb024cbf9b038b8219b861e363fe474423bccc6122c3183b4d1437`.

A fresh external f2 workspace was made as a hard-link clone of the historical
source tree, then given only the reviewed current probe and the separately
pinned lock.  The compatibility adapter is generated, not a production
change.  The reviewed baseline preparation completed with these artifacts:

| artifact | bytes | SHA-256 |
|---|---:|---|
| `f2-baseline/litchi_w8_fixture_descriptor.json` | 4,231 | `b7439d899a940b4135dd44081554a4aa9d7e70c19e2ccd3339440baad5ca7e09` |
| `f2-baseline/original-baseline-route-map.json` | 264,956 | `15a223fd3a3e99c4196b1e652582b8374e00c368f68de0196a992502850af96e` |
| `f2-baseline/original-baseline-consumer-adapter.patch` | 153,482 | `3e05d33946c7a9a2fea5c84cda33c799fbf8cd3c92fbcf890812736641ed01cd` |
| `f2-baseline/original-baseline-source-assembly.json` | 796 | `70a6e831a247d3bbd501299fa1525d7a99afa11801d1ef92ada33b43d06f4ab1` |
| `f2-baseline/litchi_w8_dependency_projection.json` | 146,063 | `b3a75410e2fe4a8805cc13e11b1759f0b7487f13c3fa8682d1cb363d771cd824` |
| `f2-baseline/original-baseline-assembly.json` | 5,795 | `03b2d62de10864ca3750fadbe16763cf161e38f50f5e36988d169278f00a494b` |

The generated adapter is 74,719 bytes, SHA-256
`3e14376c366fa911025212d7e7081ab5c3b5aaac804c1e01a1415fd5cd826ca0`.
It changes exactly the four old result-field consumers and removes exactly the
three f2 XLSB worksheet-resolution expressions.  The assembly status is
`prepared_consumer_adapter_only`; its native compile, native validation and
performance fields remain `pending`, with `acceptance_claim: false`.
These six f2 artifacts were copied byte-for-byte into the f6fe lane's
`provenance/f2-baseline` directory; the copy is not a new endpoint identity.
The copied JSON retains the original external f2 workspace/patch paths from
the preparation run.  Until a coordinator regenerates it in a different
workspace, endpoint construction must use those existing paths (or the
original f2 output directory), never rewrite the JSON just to change a path.

The exact build-free preparation command was:

```sh
C=/home/zhuhe/code/litchi-worktrees/uo2-budget-validation
LANE=(deleted-coord-artifact)
F2="$LANE/f2-workspace"
python3 "$C/tools/w8_original_baseline.py" prepare \
  --git-root "$C" \
  --workspace "$F2" \
  --probe "$F2/crates/litchi-ops/examples/file_w8_qualified_probe.rs" \
  --lock "$F2/Cargo.lock" \
  --lock-source-commit f2bbbc8d07ab5da26c180f1ff74813d39f632d54 \
  --lock-origin "pinned offline resolution generated from the current candidate's reviewed f2-compatible dependency resolution; copied as a fresh external input" \
  --source-archive (deleted-coord-artifact) \
  --output-dir "$LANE/provenance/f2-baseline" \
  --target x86_64-unknown-linux-gnu --profile release
```

The command performs no Cargo invocation.  The historical endpoint's old
probe, lock, projection, adapter, binary and validation identities remain
historical evidence and are not used as current inputs.

## Pairing prerequisite found by the strict runner

The shared fixture, probe and lock are intentionally identical, but the two
fresh projections are not:

```text
f2 compatibility workspace projection: b3a75410e2fe4a8805cc13e11b1759f0b7487f13c3fa8682d1cb363d771cd824
f6fe candidate projection:             586cbbe06a8591322488a9650bbc098ad4b9f2f137c95e2ea80617728cbcd2e5
```

This is a real manifest/dependency-graph difference, not stale hashing.  f6fe
declares the later `litchi-biff`, `litchi-ograph`, `soapberry-zip` and `sha2`
dependencies and the `siphasher` workspace dependency, and its `ppt`, `xls`
and ODF feature declarations include those dependencies.  The pristine f2
manifest does not.  `results/0362h/run.py` requires
`dependency_projection_sha256` to match between endpoints, and each endpoint
manifest recomputes its projection from its own workspace; therefore it is
not valid to attach the candidate projection to f2 or to relabel the f2
projection.  Native endpoint pairing is blocked until the coordinator reviews
an accepted compatibility strategy or chooses a final source whose manifests
share the projection.  This receipt makes no qualification claim while that
identity check is unresolved.

## Matched build-graph overlay proposal

The exact external proposal is retained at
`(deleted-coord-artifact)`:

| artifact | bytes | SHA-256 |
|---|---:|---|
| `manifest-example-overlay.patch` | 94,042 | `64e06bbcd282da4176edd7c38a9d8d806b916fa05799a0e440f020451399e5c8` |
| `manifest-example-overlay-receipt.json` | 16,606 | `1c12f7055ca4edb57fb1a583a875531b03af94c5b8fb84335fa344ac37cc0bf0` |

The allowlist is exact: root `Cargo.toml`;
`crates/litchi-iwa-core/Cargo.toml`, `crates/litchi-odraw/Cargo.toml`,
`crates/litchi-ops/Cargo.toml`, `crates/litchi-pages/Cargo.toml`,
`crates/litchi-ppt/Cargo.toml`, `crates/litchi-xlsb/Cargo.toml`, and
`crates/litchi-xlsx/Cargo.toml`; plus the three files newly named by those
manifests: `crates/litchi-pages/examples/generate_footnote_fixture.rs`,
`crates/litchi-ops/examples/xls_import_scaling.rs`, and
`crates/litchi-ops/examples/odf_shared_triad_qualification.rs`.  The overlay
contains no `crates/*/src/**` path.  Thus it changes the baseline build graph
and auxiliary manifest targets only; it does not change the f2 runtime source
implementation.  It must still be described as a compatibility-overlay
baseline, not as pristine f2.

The historical and candidate locks contain the same 586 package identities
(name, version, source and checksum), with equality for every common package.
The six local dependency-edge differences are exactly the manifest additions
for `litchi-iwa-core`, `litchi-odraw`, `litchi-ops`, `litchi-ppt`,
`litchi-xlsb`, and `litchi-xlsx`; the candidate lock is the lock both sides
would use after applying this overlay.  The all-features selection, target,
profile, workload and common resolved package identities are therefore
explicitly checkable without changing `shared_identity`.

The existing `litchi_w8_source_assembly.v1` shape can carry this as
`overlay.kind = "compatibility"` with one patch artifact, but it cannot
faithfully prove the proposal: it has no scope allowlist, runtime-source
guard, lock-edge receipt, or matched-projection binding.  The current
`w8_original_baseline.py` also intentionally rejects these manifest additions
and permits only the corrected probe, adapter and lock.  A reviewed overlay
requires a tool/schema extension (for example a v2 source-assembly overlay
with scope, runtime guard, lock identity and projection identity fields).
No tooling was changed while the f6fe gate is frozen; the proposal receipt
records this as pending review.

## Historical diagnostic recipe (strict runner; not qualification)

The following commands are recorded but intentionally not run in this lane.
They are a historical diagnostic recipe that ends at the ordinary strict
`shared_identity` ABBA path.  The f2 compatibility projection and the f6fe
projection differ, so that path must reject this pair before timing; it is not
an acceptance recipe and must not be reported as current native evidence.
The f6fe production checkout used below predates the explicit 0362yh policy and
does not contain its pair-receipt protocol.  They use independent targets and
temporary directories, one Cargo job, the 20-GB memory-capped Cargo shim,
offline/locked resolution, and release debug information disabled.  They must
be executed only after the coordinator freezes the paired source inputs and
confirms no forbidden process overlap.

```sh
LANE=(deleted-coord-artifact)
CAND=/home/zhuhe/code/litchi-worktrees/uo2-w8-f6fe-native-endpoint
BASE_LANE=(deleted-coord-artifact)
F2="$BASE_LANE/f2-workspace"
export PATH=(deleted-coord-artifact):$PATH

BASE_BUILD="env CARGO_TARGET_DIR=$LANE/target-f2 TMPDIR=$LANE/tmp/f2-build CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_NET_OFFLINE=true LITCHI_MEM_MAX=20G cargo build --offline --locked --release --jobs 1 --manifest-path $F2/Cargo.toml -p litchi-ops --example file_w8_original_baseline_probe --all-features"
CAND_BUILD="env CARGO_TARGET_DIR=$LANE/target-f6fe TMPDIR=$LANE/tmp/f6fe-build CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_NET_OFFLINE=true LITCHI_MEM_MAX=20G cargo build --offline --locked --release --jobs 1 --manifest-path $CAND/Cargo.toml -p litchi-ops --example file_w8_qualified_probe --all-features"

env CARGO_TARGET_DIR="$LANE/target-f2" TMPDIR="$LANE/tmp/f2-build" \
  CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=0 \
  CARGO_NET_OFFLINE=true LITCHI_MEM_MAX=20G \
  cargo build --offline --locked --release --jobs 1 \
  --manifest-path "$F2/Cargo.toml" -p litchi-ops \
  --example file_w8_original_baseline_probe --all-features

env CARGO_TARGET_DIR="$LANE/target-f6fe" TMPDIR="$LANE/tmp/f6fe-build" \
  CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=0 \
  CARGO_NET_OFFLINE=true LITCHI_MEM_MAX=20G \
  cargo build --offline --locked --release --jobs 1 \
  --manifest-path "$CAND/Cargo.toml" -p litchi-ops \
  --example file_w8_qualified_probe --all-features
```

After each successful build, run the complete untimed 109-route check before
any timing.  The `--source`/`--probe`/`--lock` paths below must remain the
same inputs bound above, and the build command stored in each receipt must be
the exact command used:

```sh
python3 -B "$CAND/docs/unified-ops/results/0362h/validate.py" \
  "$LANE/target-f2/release/examples/file_w8_original_baseline_probe" \
  "$LANE/validation/f2" \
  --source (deleted-coord-artifact) \
  --probe "$F2/crates/litchi-ops/examples/file_w8_qualified_probe.rs" \
  --lock "$F2/Cargo.lock" \
  --source-commit f2bbbc8d07ab5da26c180f1ff74813d39f632d54 \
  --build-command "$BASE_BUILD"

python3 -B "$CAND/docs/unified-ops/results/0362h/validate.py" \
  "$LANE/target-f6fe/release/examples/file_w8_qualified_probe" \
  "$LANE/validation/f6fe" \
  --source "$LANE/f6fe-source.tar" \
  --probe "$CAND/crates/litchi-ops/examples/file_w8_qualified_probe.rs" \
  --lock "$CAND/Cargo.lock" \
  --source-commit f6fe430a4d19d94aa360c0e9ed03ebd757265225 \
  --build-command "$CAND_BUILD"
```

Then create the two endpoint manifests.  This is still provenance checking,
not a build or timing step:

```sh
LANE=(deleted-coord-artifact)
CAND=/home/zhuhe/code/litchi-worktrees/uo2-w8-f6fe-native-endpoint
BASE_LANE=(deleted-coord-artifact)
F2="$BASE_LANE/f2-workspace"
HIST=(deleted-coord-artifact)
BASE_PROV="$BASE_LANE/provenance/f2-baseline"

python3 "$CAND/tools/w8_file_provenance.py" endpoint \
  --binary "$LANE/target-f2/release/examples/file_w8_original_baseline_probe" \
  --source "$HIST/f2-source.tar" --lock "$F2/Cargo.lock" \
  --probe "$F2/crates/litchi-ops/examples/file_w8_qualified_probe.rs" \
  --fixture "$BASE_PROV/litchi_w8_fixture_descriptor.json" \
  --dependency-projection "$BASE_PROV/litchi_w8_dependency_projection.json" \
  --workspace "$F2" --workspace-manifest "$F2/Cargo.toml" \
  --package-manifest "$F2/crates/litchi-ops/Cargo.toml" \
  --toolchain "$F2/rust-toolchain.toml" \
  --source-assembly "$BASE_PROV/original-baseline-source-assembly.json" \
  --build-receipt "$LANE/endpoint/f2/build-receipt.json" \
  --validation-receipt "$LANE/validation/f2/receipt.json" \
  --source-commit f2bbbc8d07ab5da26c180f1ff74813d39f632d54 \
  --all-features --target x86_64-unknown-linux-gnu --profile release \
  --output "$LANE/endpoint/f2/endpoint.json"

python3 "$CAND/tools/w8_file_provenance.py" endpoint \
  --binary "$LANE/target-f6fe/release/examples/file_w8_qualified_probe" \
  --source "$LANE/f6fe-source.tar" --lock "$CAND/Cargo.lock" \
  --probe "$CAND/crates/litchi-ops/examples/file_w8_qualified_probe.rs" \
  --fixture "$LANE/provenance/litchi_w8_fixture_descriptor.json" \
  --dependency-projection "$LANE/provenance/litchi_w8_dependency_projection.json" \
  --workspace "$CAND" --workspace-manifest "$CAND/Cargo.toml" \
  --package-manifest "$CAND/crates/litchi-ops/Cargo.toml" \
  --toolchain "$CAND/rust-toolchain.toml" \
  --source-assembly "$LANE/provenance/f6fe-source-assembly.json" \
  --build-receipt "$LANE/endpoint/f6fe/build-receipt.json" \
  --validation-receipt "$LANE/validation/f6fe/receipt.json" \
  --source-commit f6fe430a4d19d94aa360c0e9ed03ebd757265225 \
  --all-features --target x86_64-unknown-linux-gnu --profile release \
  --output "$LANE/endpoint/f6fe/endpoint.json"
```

Only after both complete receipts and endpoint manifests pass the strict
identity checks may the quiet-host timing command be considered:

```sh
bash "$CAND/docs/unified-ops/results/0362h/abba.sh" \
  "$LANE/target-f2/release/examples/file_w8_original_baseline_probe" \
  "$LANE/target-f6fe/release/examples/file_w8_qualified_probe" \
  "$LANE/abba" 5 29 \
  --before-manifest "$LANE/endpoint/f2/endpoint.json" \
  --after-manifest "$LANE/endpoint/f6fe/endpoint.json"
python3 "$CAND/docs/unified-ops/results/0362h/analyze.py" "$LANE/abba"
```

The ABBA command requires at least five rounds and a quiet host.  It was not
run here, and no native performance or acceptance status is implied.  Its
ordinary strict identity failure is retained only to document the historical
diagnostic boundary.

## Reviewed pair-receipt protocol (opt-in, separate checkout)

When the coordinator has regenerated endpoint manifests that match the fixed
0362yh policy pins, use a clean protocol checkout separate from the f6fe
production endpoint checkout above.  The f6fe checkout cannot run this path:
it intentionally has no policy or pair-aware runner.  The protocol checkout
below is pinned to the reviewed implementation commit
`2b1f588560c0cbf2fc7cd549944dda7601a51352`; create it detached from a
repository containing that commit and verify it is clean before use.  Do not
build from this checkout: it supplies only the policy, receipt tool and runner;
the binaries and endpoint manifests remain the separately pinned lane inputs.

```sh
PROTO=/home/zhuhe/code/litchi-worktrees/uo2-w8-native-protocol-2b1f588560
PROTO_COMMIT=2b1f588560c0cbf2fc7cd549944dda7601a51352
# One-time setup from a repository containing PROTO_COMMIT:
# git -C /path/to/protocol-repository worktree add --detach "$PROTO" "$PROTO_COMMIT"
test "$(git -C "$PROTO" rev-parse HEAD)" = "$PROTO_COMMIT"
test -z "$(git -C "$PROTO" status --porcelain)"

LANE=(deleted-coord-artifact)
POLICY="$PROTO/docs/unified-ops/results/0362yh/original-to-final-policy.json"
PAIR_RECEIPT="$LANE/endpoint/original-to-final-pair-receipt.json"

python3 "$PROTO/tools/w8_projection_policy.py" pair \
  --policy "$POLICY" \
  --before-manifest "$LANE/endpoint/f2/endpoint.json" \
  --after-manifest "$LANE/endpoint/f6fe/endpoint.json" \
  --output "$PAIR_RECEIPT"

bash "$PROTO/docs/unified-ops/results/0362h/abba.sh" \
  "$LANE/target-f2/release/examples/file_w8_original_baseline_probe" \
  "$LANE/target-f6fe/release/examples/file_w8_qualified_probe" \
  "$LANE/abba" 5 29 \
  --before-manifest "$LANE/endpoint/f2/endpoint.json" \
  --after-manifest "$LANE/endpoint/f6fe/endpoint.json" \
  --original-to-final-policy "$POLICY" \
  --original-to-final-receipt "$PAIR_RECEIPT"
python3 "$PROTO/docs/unified-ops/results/0362h/analyze.py" "$LANE/abba"
```

The pair command is fail-closed and must be run before any timing.  If these
lane manifests still carry the historical mismatched projection, it rejects
them; that rejection is evidence of stale inputs, not a native result.  A
successful receipt is protocol-only until the final production source freeze,
native validation and ABBA review are separately recorded.
