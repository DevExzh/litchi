# Paired XLSB creation benchmark

This standalone harness compares the minimal native XLSB creator with the
unified creation entry point. The native leg measures
`Package::new().into_workbook()`, which gives callers the same ready workbook
projection that the unified creator adopts. The unified leg measures
`litchi_ops::create(FileFormat::Xlsb)`, returning a unified `File`.

The default run uses nine ABBA rounds and 18 samples per leg. A warm-up creates
one owner through each route, verifies that the caller can read the sole
worksheet named `Sheet1`, serializes both outputs, and compares the exact ZIP
bytes. If container bytes differ, every decoded ZIP member is compared and the
report identifies any differing member by byte length and SHA-256. The warm-up
and all output serialization and validation are outside the timed interval.

This benchmark measures the initial package plus ready projection only. It
makes no open or write performance claim and does not include opening a
caller-provided file or writing to a caller-provided sink.

Run from the repository root after the native `Package::new` and unified XLSB
creator changes are present, with no concurrent build, test, or measurement
job:

```sh
BUILD=/tmp/litchi-unified-xlsb-create-target
BIN="$BUILD/release/unified-xlsb-create-baseline"
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR="$BUILD" cargo build --release --locked \
  --manifest-path tools/perf-baseline/unified-xlsb-create/Cargo.toml
LITCHI_SOURCE_COMMIT="$(git rev-parse HEAD)" \
LITCHI_SOURCE_DIGEST="$(git diff HEAD --binary | sha256sum | awk '{print $1}')" \
LITCHI_RUSTC_VERSION="$(rustc --version)" \
LITCHI_TARGET="$(rustc -vV | sed -n 's/^host: //p')" \
LITCHI_MACHINE="$(hostname)" \
LITCHI_BINARY_SHA256="$(sha256sum "$BIN" | awk '{print $1}')" \
LITCHI_BUILD_COMMAND='CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=/tmp/litchi-unified-xlsb-create-target cargo build --release --locked --manifest-path tools/perf-baseline/unified-xlsb-create/Cargo.toml' \
taskset -c 0 "$BIN" > /tmp/unified-xlsb-create.json
```

Use `--rounds N` for a narrower smoke measurement; at least three ABBA rounds
are required. The JSON records raw samples, upper-middle medians, the ratio,
warm-up output hashes and equivalence, the caller-readiness check, failure
matrix scope, and source/build/machine provenance. Never commit the compiled
binary. Keep larger raw captures outside the checkout and keep any reduced
evidence below the repository size limits.
