# 0956 bounded reproduction commands

These commands document regeneration only. They were not run while the formal
0956 capture was active. Use fresh `/tmp` destinations and retain the original
external packet and executables.

## Source tree and manifest

The frozen base is Git commit `c8e80705989183d88950e282ef1852f3e3470b10`.
The source freeze excludes `docs/performance/**` and `docs/report/**`; five
approved files are overlaid from the retained external `source-v1` snapshot:

```sh
set -eu
REPO=/home/zhuhe/code/litchi
BASE_COMMIT=c8e80705989183d88950e282ef1852f3e3470b10
FROZEN=/home/zhuhe/litchi-performance-0956/source-v1
REBUILD=/tmp/litchi-0956-source-v3
if test -e "$REBUILD"; then
  echo "refusing existing regeneration directory: $REBUILD" >&2
  exit 2
fi
mkdir "$REBUILD"
git -C "$REPO" archive --format=tar "$BASE_COMMIT" -- . \
  ':(exclude)docs/performance/**' ':(exclude)docs/report/**' |
  tar -xf - -C "$REBUILD"
for rel in \
  tools/perf-baseline/README.md \
  tools/perf-baseline/src/lib.rs \
  tools/perf-baseline/src/xlsx_managed_compression.rs \
  tools/perf-baseline/src/xlsx_managed_compression_benchmark.rs \
  tools/perf-baseline/examples/xlsx_managed_compression_benchmark.rs; do
  mkdir -p "$REBUILD/$(dirname "$rel")"
  cp --preserve=mode,timestamps "$FROZEN/$rel" "$REBUILD/$rel"
done
chmod -R g+rwX "$REBUILD"
```

Generate a new manifest path with the bounded helper copied in this packet:

```sh
python3 -B docs/performance/results/change-0956/tools/reproduce_source_manifest_0956.py \
  --source-root "$REBUILD" \
  --output /tmp/source-manifest-v3-regenerated.json \
  --base-commit "$BASE_COMMIT" \
  --measurement-source-commit "$BASE_COMMIT" \
  --base-source-root source-base-v1 \
  --base-manifest-sha256 e5c8a42d59ea94d560460d4e42c54f42c73918acef701bd1f4266797c1735273 \
  --prior-integration-manifest-sha256 f044b21d477113b0c729230fb5b12d13c315259069adc43dd607b4ba715db49f \
  --integration-revision approved-overlay-v3 \
  --overlay-path tools/perf-baseline/README.md \
  --overlay-path tools/perf-baseline/src/lib.rs \
  --overlay-path tools/perf-baseline/src/xlsx_managed_compression.rs \
  --overlay-path tools/perf-baseline/src/xlsx_managed_compression_benchmark.rs \
  --overlay-path tools/perf-baseline/examples/xlsx_managed_compression_benchmark.rs
sha256sum /tmp/source-manifest-v3-regenerated.json
```

The frozen external manifest is
`/home/zhuhe/litchi-performance-0956/source-manifest-v3.json`, 2,266,359
bytes, SHA-256
`fd1bc21ee257a1e9ad3b68b46be7cdbfd677877cffa17b714e3493de96b18193`. It
contains 10,930 files totaling 273,393,096 source bytes. `cmp` against the
regenerated file is the exact identity check; the command must not silently
replace the frozen manifest. The helper records file bytes, SHA-256, and modes;
its output is a new artifact rather than a claim that relocated paths are the
original source freeze.

## Release binaries

The recorded build command is:

```sh
mkdir -p /tmp/litchi-0956-build-v1/tmp/integration-native-v1
export CARGO_HOME=/home/zhuhe/.cargo RUSTUP_HOME=/home/zhuhe/.rustup
export CARGO_TARGET_DIR=/tmp/litchi-0956-build-v1/target-base-lib-v1
export CARGO_NET_OFFLINE=true CARGO_INCREMENTAL=0 CARGO_BUILD_JOBS=2
export CARGO_PROFILE_RELEASE_CODEGEN_UNITS=1 CARGO_PROFILE_RELEASE_DEBUG=1
export CARGO_PROFILE_RELEASE_LTO=fat CARGO_PROFILE_RELEASE_OPT_LEVEL=3
export RUSTUP_TOOLCHAIN=1.95.0 LANG=C LC_ALL=C TZ=UTC RUST_BACKTRACE=0
export TMPDIR=/tmp/litchi-0956-build-v1/tmp/integration-native-v1
export TEMP="$TMPDIR" TMP="$TMPDIR"
unset RUSTFLAGS CARGO_ENCODED_RUSTFLAGS RUSTC_WRAPPER CARGO_BUILD_RUSTC_WRAPPER
cargo +1.95.0 build --release --locked --offline \
  --manifest-path /tmp/litchi-0956-source-v3/tools/perf-baseline/Cargo.toml \
  --example xlsx_managed_compression \
  --example xlsx_managed_compression_benchmark -j2
sha256sum \
  "$CARGO_TARGET_DIR/release/examples/xlsx_managed_compression" \
  "$CARGO_TARGET_DIR/release/examples/xlsx_managed_compression_benchmark"
```

The retained external executable identities are:

- `xlsx_managed_compression`: 43,172,232 bytes,
  SHA-256 `73d0c9863cc16478e17e1b09402a42f45c3dceba9e24f4a853187d983e32e75a`.
- `xlsx_managed_compression_benchmark`: 43,250,464 bytes,
  SHA-256 `38381c594b6dd5924e07d3a27f57bf15729811b8f72d7353aac5aac810741304`.

These hashes were produced at the recorded external source/target paths. A
relocated build regenerates the same executable roles but can contain different
debug-path bytes; retain its output receipt and do not overwrite the active
external binaries or run this command during a live capture.
