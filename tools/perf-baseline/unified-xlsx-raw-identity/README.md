# XLSX raw identity lifecycle probe

This probe records the immediate pre-repair pair
`0befb112d96d2409e6f40c9abfeb4813665905ad` and
`5726f03c7db2dd7dd539000b45f109f14ea8c33a`, then carries the matched
lifecycle into the current assembly. The versioned OPC ingress is deliberate:
the historical source uses its one-argument `OpcPackage::from_vec` default,
while the current source uses the required-limit
`OpcPackage::from_vec(bytes, SourceLimits::default())`. `--api-variant` labels
the report; Cargo features select the source-compatible constructor.

The current v2 harness is based on `a3b649ca857673c9c0ce62e04c7d44e7ba78c502`.
After the adapter change its source-file SHA-256 is
`3c9d07e449865ddeee86ec273cf2ec31eaccb9897f5631bc6db45416b56f14fa`; the
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
its one binary includes the existing counting allocator wrapper and dispatches
only the `xlsx-raw-identity` selector. It therefore does not compile the
full `tools/perf-baseline` package or its unrelated DOCX macro targets.

The standalone manifest intentionally has no lockfile until the coordinator
supplies the final v9 candidate revision. Generate the lockfile in the exact
source overlay immediately before the first build, record its SHA-256 beside
the source and binary identities, and then use `--locked` for every measured
build. No result is qualified until that lockfile and the final candidate
revision are recorded.

The coordinator must keep captures outside the checkout and reuse existing
target directories. The recorded directories for this pair are
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets/reader-paths-e63`
for the pinned historical source and
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets/hist-current` for the
current assembly. At preparation time they are 1.3 GiB and 1.6 GiB,
respectively. The existing historical `release/xlsx_source_reader_paths_alloc`
binary is retained; the commands below use the new
`release/xlsx-raw-identity-alloc` name and abort if that path is already
occupied. Do not clean either target or overwrite any recorded binary.

The feature selection is part of the reproducibility contract. In the
standalone manifest, `default = ["current-xlsx-raw-opc"]`, so the ordinary
default build and `--all-features` both select the required-limit constructor;
when both current and legacy markers are present, the current cfg wins. A
`--no-default-features` build without either marker also uses the current
required-limit constructor. Only `--no-default-features --features
allocator-metrics,legacy-xlsx-raw-opc` selects the historical one-argument
constructor.

The historical checkout is an external source overlay, not a modified
full-workspace build. The overlay copies the exact standalone manifest and
the two probe modules from the final candidate into an archive of the pinned
`0befb112d96d2409e6f40c9abfeb4813665905ad` tree. This keeps the probe source
identical while avoiding the unrelated DOCX macro failures. Root must replace
`<FINAL_CANDIDATE_REVISION>` with the post-v9 final assembly hash; no earlier
raw branch or 5726-only result is an acceptable substitute:

```sh
HIST_TARGET=/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets/reader-paths-e63
CURRENT_TARGET=/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets/hist-current
FINAL_CHECKOUT=/path/to/final-assembly-checkout
FINAL_CANDIDATE_REVISION=<FINAL_CANDIDATE_REVISION>
BASELINE_REVISION=0befb112d96d2409e6f40c9abfeb4813665905ad
EVIDENCE=/path/to/coordinator-evidence/raw-identity

test ! -e "$HIST_TARGET/release/xlsx-raw-identity-alloc"
test ! -e "$CURRENT_TARGET/release/xlsx-raw-identity-alloc"
mkdir -p "$EVIDENCE"

# Make a disposable baseline source overlay. This does not touch the pinned
# checkout, the final checkout, or any format source. The final candidate
# supplies the standalone manifest, allocation module, and identical v2 probe.
BASELINE_OVERLAY=$(mktemp -d /tmp/litchi-xlsx-raw-baseline.XXXXXX)
git -C "$FINAL_CHECKOUT" archive --format=tar "$BASELINE_REVISION" \
  | tar -xf - -C "$BASELINE_OVERLAY"
git -C "$FINAL_CHECKOUT" archive --format=tar "$FINAL_CANDIDATE_REVISION" \
  tools/perf-baseline/unified-xlsx-raw-identity \
  tools/perf-baseline/src/allocation_metrics.rs \
  tools/perf-baseline/src/xlsx_raw_identity.rs \
  | tar -xf - -C "$BASELINE_OVERLAY"

# The overlay's manifest has path dependencies into its archived crates.
# Generate its lockfile once for this exact source pair; record the resulting
# Cargo.lock SHA-256 before using --locked below.
source /home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/env.sh
CARGO_BUILD_JOBS=1 CARGO_NET_OFFLINE=true cargo generate-lockfile \
  --manifest-path "$BASELINE_OVERLAY/tools/perf-baseline/unified-xlsx-raw-identity/Cargo.toml"
sha256sum "$BASELINE_OVERLAY/tools/perf-baseline/unified-xlsx-raw-identity/Cargo.lock" \
  > "$EVIDENCE/raw-baseline-lock.sha256"

CARGO_BUILD_JOBS=1 CARGO_NET_OFFLINE=true cargo build --release --locked \
  --no-default-features --features allocator-metrics,legacy-xlsx-raw-opc \
  --manifest-path "$BASELINE_OVERLAY/tools/perf-baseline/unified-xlsx-raw-identity/Cargo.toml" \
  --bin xlsx-raw-identity-alloc --target-dir "$HIST_TARGET"

sha256sum "$HIST_TARGET/release/xlsx-raw-identity-alloc" \
  > "$EVIDENCE/raw-baseline-binary.sha256"
taskset -c 0 "$HIST_TARGET/release/xlsx-raw-identity-alloc" \
  xlsx-raw-identity --api-variant original \
  --source-revision "$BASELINE_REVISION" \
  --iterations 5 --warmups 2 --json "$EVIDENCE/raw-original.json"

# Build the final assembly through its thin manifest. The default command
# proves the required-limit adapter; --all-features proves current wins over
# legacy. Both builds use one job and the same existing target directory.
CURRENT_MANIFEST="$FINAL_CHECKOUT/tools/perf-baseline/unified-xlsx-raw-identity/Cargo.toml"
CARGO_BUILD_JOBS=1 CARGO_NET_OFFLINE=true cargo generate-lockfile \
  --manifest-path "$CURRENT_MANIFEST"
sha256sum "$FINAL_CHECKOUT/tools/perf-baseline/unified-xlsx-raw-identity/Cargo.lock" \
  > "$EVIDENCE/raw-current-lock.sha256"
CARGO_BUILD_JOBS=1 CARGO_NET_OFFLINE=true cargo build --release --locked \
  --features allocator-metrics --manifest-path "$CURRENT_MANIFEST" \
  --bin xlsx-raw-identity-alloc --target-dir "$CURRENT_TARGET"
sha256sum "$CURRENT_TARGET/release/xlsx-raw-identity-alloc" \
  > "$EVIDENCE/raw-current-default-binary.sha256"

CARGO_BUILD_JOBS=1 CARGO_NET_OFFLINE=true cargo build --release --locked \
  --all-features --manifest-path "$CURRENT_MANIFEST" \
  --bin xlsx-raw-identity-alloc --target-dir "$CURRENT_TARGET"
sha256sum "$CURRENT_TARGET/release/xlsx-raw-identity-alloc" \
  > "$EVIDENCE/raw-current-all-features-binary.sha256"

taskset -c 0 "$CURRENT_TARGET/release/xlsx-raw-identity-alloc" \
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
