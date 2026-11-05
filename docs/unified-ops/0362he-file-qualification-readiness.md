# 0362he — W8 File provenance and baseline readiness

- **Scope:** deterministic fixture/dependency provenance and baseline selection
  only; no Cargo build, timing run or native qualification is performed here.
- **Status:** the corrected 109-route source is ready for provenance
  generation.  Quiet-host native timing remains pending.  The diagnostic
  `0359qe` rows do not change that boundary.

## What is now reproducible

`tools/w8_file_provenance.py` is the single source of truth for the inputs that
the guarded runner consumes.  It reads the checked-in
`crates/litchi-ops/examples/file_w8_qualified_probe.rs`; it does not invent an
empty fixture file or accept a hand-written route list.

The `fixture` command checks the probe's family order, route construction,
fixture constructor call sites, `PARAGRAPHS`/`SHAPES`/calibration constants,
current `Commit::into_file()` extraction and list/validate modes.  It writes a
canonical `litchi_w8_fixture_descriptor.v1` containing the probe SHA-256, the
11 family recipes and all 109 routes.  A changed constructor, count, batch
bound or route guard fails descriptor generation until the recipe is reviewed.

The `dependency` command parses the supplied `Cargo.lock`, workspace and
`litchi-ops` manifests, selected feature set, target, release profile and
`rust-toolchain.toml`.  It re-generates the fixture descriptor from the
reviewed probe and compares the complete canonical descriptor, rather than
accepting a schema plus a probe hash.  The resulting
`litchi_w8_dependency_projection.v1` contains the resolved package graph and
hashes of every source input; it does not run Cargo.  The `source` command
records the full resolved source commit and, where needed, an explicitly
described compatibility overlay and patch artifact.  The `endpoint` command
then binds the executable, source assembly, lock, probe, fixture, projection,
workspace/package manifests, toolchain, build receipt and complete untimed
109-route validation receipt into the `litchi_file_endpoint.v1` manifest
expected by `results/0362h/run.py`.
The package declaration must be `litchi-ops` at the canonical
`crates/litchi-ops/Cargo.toml` path, and the probe must be at
`crates/litchi-ops/examples/file_w8_qualified_probe.rs` in that same
workspace.  The runner rechecks both path and package-name identities before
calibration, resolves `version.workspace = true` from the workspace package,
and compares that version with the projected `litchi-ops` lock entry.

Keep all generated descriptors, lockfiles, source archives, manifests and
binaries in an external scratch directory.  The tool rejects
`3rdparty/`, `docs/performance/results/` and `docs/report/` as input/output
locations where applicable.  The repository contains only this recipe and the
commands below, not generated qualification artifacts.

Example preparation (run after the final W4/W4b-integrated source is chosen;
these commands intentionally do not build):

```sh
set -eu
SRC=/path/to/source-checkout
PROBE="$SRC/crates/litchi-ops/examples/file_w8_qualified_probe.rs"
LOCK=/path/to/locked/Cargo.lock
OUT=(deleted-coord-artifact)
SOURCE_ARCHIVE=/path/to/source.tar
BASE=/path/to/historical-baseline-checkout
OVERLAY_BASE="$(git -C "$BASE" rev-parse HEAD)"
OVERLAY_PATCH=/path/to/reviewed-compatibility.patch
mkdir -p "$OUT"

python3 tools/w8_file_provenance.py fixture \
  --probe "$PROBE" --output "$OUT/fixture.json"
python3 tools/w8_file_provenance.py dependency \
  --workspace "$SRC" \
  --package-manifest "$SRC/crates/litchi-ops/Cargo.toml" \
  --lock "$LOCK" --toolchain "$SRC/rust-toolchain.toml" \
  --probe "$PROBE" --fixture "$OUT/fixture.json" \
  --all-features --target x86_64-unknown-linux-gnu --profile release \
  --output "$OUT/dependency.json"

python3 tools/w8_file_provenance.py source \
  --source "$SOURCE_ARCHIVE" \
  --source-commit "$(git -C "$SRC" rev-parse HEAD)" \
  --overlay-kind compatibility --overlay-base-commit "$OVERLAY_BASE" \
  --overlay-patch "$OVERLAY_PATCH" \
  --description "audited compatibility overlay rooted at the selected baseline" \
  --output "$OUT/source-assembly.json"

python3 tools/w8_file_provenance.py endpoint \
  --binary /path/to/protected/endpoint \
  --source "$SOURCE_ARCHIVE" \
  --lock "$LOCK" --probe "$PROBE" --fixture "$OUT/fixture.json" \
  --dependency-projection "$OUT/dependency.json" \
  --workspace "$SRC" --workspace-manifest "$SRC/Cargo.toml" \
  --package-manifest "$SRC/crates/litchi-ops/Cargo.toml" \
  --toolchain "$SRC/rust-toolchain.toml" \
  --source-assembly "$OUT/source-assembly.json" \
  --build-receipt /path/to/external/build-receipt.json \
  --validation-receipt /path/to/external/untimed-validation.json \
  --source-commit "$(git -C "$SRC" rev-parse HEAD)" \
  --all-features --target x86_64-unknown-linux-gnu --profile release \
  --output "$OUT/endpoint.json"
```

Generate the fixture and dependency projection once for the shared workload
and use the resulting identities in both endpoint manifests.  The runner
requires the lock, projection, probe, fixture, feature selection, target,
profile and dependency/source metadata to be internally consistent on each
endpoint, and requires the shared lock/projection/probe/fixture identities to
match across endpoints.  It rechecks the complete fixture/projection shape,
full source commit, overlay artifact and build/compiler receipt before any
route calibration.  It only starts timing after both endpoint route lists
match the complete 109-route fixture; a nonempty or partial route list is not
qualification evidence.

## Baseline decision before an expensive build

The original W8 script is explicit: `docs/unified-ops/results/0362/abba.sh`
names `f2bbbc8d07` as the trunk **before** leg and compares it with the W8
head.  That is the historical goal baseline.  It does not contain either
`file_w8_probe.rs` or the corrected probe, so a reproducible baseline needs an
external workload adapter.  The first W8 probe was introduced at
`b1ae817dc4` and has SHA-256
`b96ce20fcc35f19f3e85aacdfaea8d95dc3e9f8f5237928e1ef06541565c8605` (12,363
bytes).  It uses the pre-migration common calls (`open_bytes`, `get`, `set`,
`add`, `batch`, `import`, `merge`) and can be copied into a baseline checkout
without changing production source.  That adapter reproduces the historical
five-route workload only; it has no separate publication route, exact source
owner check or equivalent 109-route leaf/facade pair.  It is therefore
historical context, not final W8 acceptance evidence.

The corrected probe is a different source artifact: its SHA-256 is
`3a01dd2cbb971b4eba54e2eccffa99e1b3c4557db14af03e61869f99f0a0cc27` (74,747
bytes), with 109 routes and the current `Commit::into_file()` API.  It is
present from the W8 qualification commits onward, including the changed
integrated source at `d39d179353`.  `d39d179353` is not an original baseline:
it contains the equivalent-native-workload and exact-source-check repair and
must not be silently selected as the before leg.  The current assembly at
`2a98669821` carries that same corrected probe, but W4 native shared-cell work
is still under repair, so it is not a final native candidate either.

Before building, coordinator review must choose and record one of these
truthful paths:

| path | what it proves | disposition |
|---|---|---|
| `f2bbbc8d07` + the external original `file_w8_probe` adapter | historical five-route before/after context | may reproduce the old diagnostic only; not the 109-route qualification |
| an audited compatibility overlay rooted at `f2bbbc8d07`, with the unchanged corrected probe and its current API adapter | the original goal baseline on the corrected workload | required if the final native comparison is meant to claim the original W8 baseline; record the overlay patch hash and prove the same fixture/probe/projection identities |
| `d39d179353` or any later changed integration as “before” | a changed-source comparison | forbidden as the original baseline; retain only as candidate/untimed validation evidence |

Until the second path has an actual source overlay and compile proof, no native
build or ABBA run should be started.  The provenance generator and the
guarded runner do not choose a baseline implicitly.
