# XLSX raw identity lifecycle probe

This probe records the original production source
`0befb112d96d2409e6f40c9abfeb4813665905ad` and a final assembled candidate,
then carries the matched lifecycle into the current assembly. The versioned OPC ingress is deliberate:
the historical source uses its one-argument `OpcPackage::from_vec` default,
while the current source uses the required-limit
`OpcPackage::from_vec(bytes, SourceLimits::default())`. `--api-variant` labels
the report; Cargo features select the source-compatible constructor.

The current v2 harness is based on `a3b649ca857673c9c0ce62e04c7d44e7ba78c502`.
After the cross-revision relationship-target fix its source-file SHA-256 is
`a251f1199b5ba442dfd9664806b81f29c7832d071df66fc2d2d5f2ada6c97fd4`; the
adapter block (the `open_opc` documentation, fallback, and explicit
legacy-only branch) is
`294827709997633b1c0c8568e07f7a3eae4f13103a0a9757efee236c641412d5`.
The native v1 source files at both historical revisions have SHA-256
`6f5e5606afa342f87a7c1e93538b7302878a6b45113cef7b6514f62a7f1cdb34`.
Do not reuse a 5726-only result as a performance qualification for the later
current assembly.

Each corpus contains a small and a sixteen-part 512 KiB multipart package,
with one-cache and eight-cache variants. For every corpus it measures owned
open, the first selected deferred read, raw slicer insertion, raw timeline
batch insertion, typed slicer replacement, and typed timeline replacement.
The correctness pass outside the measured region checks the serialized output
bytes, native relationship IDs and paths, untouched worksheet bytes, patch
replay, inverse restoration, stale-patch refusal, no-op preservation, and
deferred decode counters.

The qualification uses the thin standalone manifest in this directory. It
depends only on `litchi-xlsx`, `litchi-opc`, and the serde/hash crates needed by
the probe. Its library imports the existing allocation and raw-probe modules;
its allocator binary includes the existing counting allocator wrapper and its
native binary dispatches the same `xlsx-raw-identity` selector without that
wrapper. It therefore does not compile the full `tools/perf-baseline` package
or its unrelated DOCX macro targets.

The standalone lockfile is committed at
`0e24ccbe5c` with SHA-256
`8b61b3c5e17b1c32d89139c97a26145117794afae83dd8fef3c9a7aff8765ee6`.
Copy that exact lockfile into the reviewed adapter in the detached baseline
worktree and use `--locked` for every measured build. No result is qualified
until the final candidate revision is also recorded in the preparation
receipt.

The preparation helper is the executable provenance gate for every consumed
adapter input; it rejects any hash other than these reviewed values, and the
runner/validator re-check the resulting receipt:

| consumed input | SHA-256 |
| --- | --- |
| `tools/perf-baseline/unified-xlsx-raw-identity/Cargo.toml` | `30ff248844c9e21436fc3913baa4c4d363daffdc617e5e7d03404423410d5c70` |
| `tools/perf-baseline/unified-xlsx-raw-identity/src/lib.rs` | `d65ce0630a50d1b972a0b2f93044091d1a5c50df1e9c8ef06e9c3ba65cf8a36a` |
| `tools/perf-baseline/unified-xlsx-raw-identity/src/bin/xlsx_raw_identity_alloc.rs` | `c21cc9b947539078885ff099527ed8d291d631b39a27bf8a299c6553add9007c` |
| `tools/perf-baseline/unified-xlsx-raw-identity/src/bin/xlsx_raw_identity_native.rs` | `71435d535e65996f2c62e9952ab3f595b8e3bd2342afc0f45092e8a407fc9ee9` |
| `tools/perf-baseline/src/allocation_metrics.rs` | `0fcec3e5c4972031b23542fd359a38401e25f32a133c77ef4c6e75facf0710d5` |
| `tools/perf-baseline/src/xlsx_raw_identity.rs` | `a251f1199b5ba442dfd9664806b81f29c7832d071df66fc2d2d5f2ada6c97fd4` |
| `tools/perf-baseline/src/bin/support/counting_allocator.rs` | `80977e55c2213eeee2af9a3133295d70cff38505e663a68e4757bf7939c60068` |
| `tools/perf-baseline/unified-xlsx-raw-identity/Cargo.lock` | `8b61b3c5e17b1c32d89139c97a26145117794afae83dd8fef3c9a7aff8765ee6` |

The coordinator must keep captures, binaries, and targets outside the
checkout. Every target used by the guarded runner must be absent or empty at
the start of the build phase; the runner never reuses, cleans, or overwrites a
target or recorded binary. Select a fresh target root for every attempt.

The feature selection is part of the reproducibility contract. In the
standalone manifest, `default = ["current-xlsx-raw-opc"]`, so the ordinary
default build and `--all-features` both select the required-limit constructor;
when both current and legacy markers are present, the current cfg wins. A
`--no-default-features` build without either marker also uses the current
required-limit constructor. Only `--no-default-features --features
allocator-metrics,legacy-xlsx-raw-opc` selects the historical one-argument
constructor.

The historical checkout is an actual detached Git worktree, not a tar archive
or directory without `.git`: the guarded runner uses Git HEAD/tree/status
identity. Create it from the original production revision, create a clean
detached candidate worktree at the selected final revision, and let the
preparation command copy only the reviewed consumer-side adapter. Its receipt
pins both trees, the exact allowed adapter diff, fixed reviewed
lock/probe/allocator hashes, and the
candidate revision. No earlier raw branch or 5726-only result is an acceptable
substitute:

```sh
HIST_TARGET=/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets/reader-paths-e63
CURRENT_DEFAULT_TARGET=/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets/current-default
CURRENT_ALL_FEATURES_TARGET=/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets/current-all-features
FINAL_CHECKOUT=/path/to/final-assembly-checkout
FINAL_CANDIDATE_REVISION=<FINAL_CANDIDATE_REVISION>
BASELINE_REVISION=0befb112d96d2409e6f40c9abfeb4813665905ad
EVIDENCE=/path/to/coordinator-evidence/raw-identity
BASELINE_SOURCE=/external/sources/xlsx-original-0bef-worktree
CANDIDATE_SOURCE=/external/sources/xlsx-final-candidate
PREPARATION_RECEIPT=/external/xlsx-raw-native/preparation.json

test ! -e "$HIST_TARGET"
test ! -e "$CURRENT_DEFAULT_TARGET"
test ! -e "$CURRENT_ALL_FEATURES_TARGET"
mkdir -p "$EVIDENCE"

git -C "$FINAL_CHECKOUT" worktree add --detach "$BASELINE_SOURCE" "$BASELINE_REVISION"
git -C "$FINAL_CHECKOUT" worktree add --detach "$CANDIDATE_SOURCE" "$FINAL_CANDIDATE_REVISION"
python3 "$FINAL_CHECKOUT/tools/native_qualification/prepare_xlsx_raw_inputs.py" \
  --baseline-root "$BASELINE_SOURCE" \
  --candidate-root "$CANDIDATE_SOURCE" \
  --candidate-revision "$FINAL_CANDIDATE_REVISION" \
  --output "$PREPARATION_RECEIPT"

# Optional evidence echo; the preparation helper above is the executable
# fixed-hash gate and this manual check is not a substitute for it.
source /home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/env.sh
sha256sum "$BASELINE_SOURCE/tools/perf-baseline/unified-xlsx-raw-identity/Cargo.lock" \
  > "$EVIDENCE/raw-baseline-lock.sha256"
test "$(awk '{print $1}' "$EVIDENCE/raw-baseline-lock.sha256")" = \
  8b61b3c5e17b1c32d89139c97a26145117794afae83dd8fef3c9a7aff8765ee6

CARGO_BUILD_JOBS=1 CARGO_NET_OFFLINE=true cargo build --release --locked \
  --no-default-features --features allocator-metrics,legacy-xlsx-raw-opc \
  --manifest-path "$BASELINE_SOURCE/tools/perf-baseline/unified-xlsx-raw-identity/Cargo.toml" \
  --bin xlsx-raw-identity-alloc --bin xlsx-raw-identity-native --target-dir "$HIST_TARGET"

sha256sum "$HIST_TARGET/release/xlsx-raw-identity-alloc" \
  > "$EVIDENCE/raw-baseline-binary.sha256"
taskset -c 0 "$HIST_TARGET/release/xlsx-raw-identity-alloc" \
  xlsx-raw-identity --api-variant original \
  --source-revision "$BASELINE_REVISION" \
  --iterations 5 --warmups 2 --json "$EVIDENCE/raw-original.json"

sha256sum "$HIST_TARGET/release/xlsx-raw-identity-native" \
  > "$EVIDENCE/raw-baseline-native-binary.sha256"
taskset -c 0 "$HIST_TARGET/release/xlsx-raw-identity-native" \
  xlsx-raw-identity --api-variant original \
  --source-revision "$BASELINE_REVISION" \
  --iterations 5 --warmups 2 --json "$EVIDENCE/raw-original-native.json"

# Callgrind is a quiet native run: keep one warmup, one measured iteration,
# and the capture outside the checkout. The full JSON still runs every case.
taskset -c 0 valgrind --tool=callgrind --quiet \
  --callgrind-out-file="$EVIDENCE/callgrind-0bef-native.out" \
  "$HIST_TARGET/release/xlsx-raw-identity-native" xlsx-raw-identity \
  --api-variant original --source-revision "$BASELINE_REVISION" \
  --iterations 1 --warmups 1 \
  --json "$EVIDENCE/raw-original-callgrind.json"

# Build the final assembly through its thin manifest. The default command
# proves the required-limit adapter; --all-features proves current wins over
# legacy. Each build gets a fresh target directory; no target is reused across
# feature routes.
CURRENT_MANIFEST="$FINAL_CHECKOUT/tools/perf-baseline/unified-xlsx-raw-identity/Cargo.toml"
sha256sum "$FINAL_CHECKOUT/tools/perf-baseline/unified-xlsx-raw-identity/Cargo.lock" \
  > "$EVIDENCE/raw-current-lock.sha256"
test "$(awk '{print $1}' "$EVIDENCE/raw-current-lock.sha256")" = \
  8b61b3c5e17b1c32d89139c97a26145117794afae83dd8fef3c9a7aff8765ee6
CARGO_BUILD_JOBS=1 CARGO_NET_OFFLINE=true cargo build --release --locked \
  --features allocator-metrics --manifest-path "$CURRENT_MANIFEST" \
  --bin xlsx-raw-identity-alloc --bin xlsx-raw-identity-native --target-dir "$CURRENT_DEFAULT_TARGET"
sha256sum "$CURRENT_DEFAULT_TARGET/release/xlsx-raw-identity-alloc" \
  > "$EVIDENCE/raw-current-default-binary.sha256"

CARGO_BUILD_JOBS=1 CARGO_NET_OFFLINE=true cargo build --release --locked \
  --all-features --manifest-path "$CURRENT_MANIFEST" \
  --bin xlsx-raw-identity-alloc --bin xlsx-raw-identity-native --target-dir "$CURRENT_ALL_FEATURES_TARGET"
sha256sum "$CURRENT_ALL_FEATURES_TARGET/release/xlsx-raw-identity-alloc" \
  > "$EVIDENCE/raw-current-all-features-binary.sha256"

taskset -c 0 "$CURRENT_DEFAULT_TARGET/release/xlsx-raw-identity-alloc" \
  xlsx-raw-identity --api-variant current \
  --source-revision "$FINAL_CANDIDATE_REVISION" \
  --iterations 5 --warmups 2 --json "$EVIDENCE/raw-current.json"
```

Run the same selected case under Callgrind for instruction counts, using the
same target, CPU pinning, iterations, warmups, and an evidence path outside
the repository. Record each binary SHA-256 from the JSON report and keep raw
Callgrind captures outside the checkout. No performance conclusion is made by
this source-only commit; the owned-open and first-read costs include the new
per-deferred-part metadata already present in the candidate and require a
matched run before qualification.
