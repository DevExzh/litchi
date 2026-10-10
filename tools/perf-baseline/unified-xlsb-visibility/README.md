# Paired XLSB sheet-visibility benchmark

This standalone harness compares one native XLSB structural edit with the
unified `File::set` route for hiding a fanout of worksheet tabs. Each generated
corpus has `N` ordinary target worksheets and one visible `VisibleSentinel`
tab. The operation hides every target tab and must preserve the sentinel as
the only visible tab. The native leg creates one `Package::edit`, stages `N`
native visibility setters, and commits once. The unified leg makes one
`File::set` call against the pre-parsed scoped query
`sheet[name!=VisibleSentinel]` with `visibility="hidden"`; the optional
`--property hidden` route uses `hidden=true` for the same semantic operation.

The default matrix measures `N=1,10,100,1000,10000` for eager and source owners,
with nine ABBA rounds. Three rounds are the minimum. All targets, query objects, and property sets are prepared before
timing. The source-native leg materializes its prepared
`SourceBackedWorkbook` into a package inside each timed edit, matching the
source facade's edit boundary.

Run from the repository root with no concurrent build, test or measurement
job:

```sh
BUILD=/tmp/litchi-unified-xlsb-visibility-target
BIN="$BUILD/release/unified-xlsb-visibility-baseline"
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR="$BUILD" cargo build --release --locked \
  --manifest-path tools/perf-baseline/unified-xlsb-visibility/Cargo.toml
LITCHI_SOURCE_COMMIT="$(git rev-parse HEAD)" \
LITCHI_SOURCE_DIGEST="$(git diff HEAD --binary | sha256sum | awk '{print $1}')" \
LITCHI_RUSTC_VERSION="$(rustc --version)" \
LITCHI_TARGET="$(rustc -vV | sed -n 's/^host: //p')" \
LITCHI_MACHINE="$(hostname)" \
LITCHI_BINARY_SHA256="$(sha256sum "$BIN" | awk '{print $1}')" \
LITCHI_BUILD_COMMAND='CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=/tmp/litchi-unified-xlsb-visibility-target cargo build --release --manifest-path tools/perf-baseline/unified-xlsb-visibility/Cargo.toml' \
taskset -c 0 "$BIN" --rounds 9 > /tmp/unified-xlsb-visibility.json
```

Override the matrix with:

```sh
unified-xlsb-visibility-baseline --counts 1,1000 --owners eager,source --rounds 3
```

The JSON records the generated corpus byte count and SHA-256, source commit,
compiler/target/build command, benchmark binary SHA-256, raw ABBA nanosecond
samples, upper-middle medians, ratios, native setter/commit accounting, and
the complete warm-up equivalence and property-alias equivalence reports. Exact ZIP bytes are preferred. If
they differ, every decoded ZIP member is compared and the report names each
missing or differing member with its byte length and SHA-256; a mismatch exits
nonzero. The output also verifies every target is hidden and the sentinel is
visible after each route. This is scoped measurement evidence, not a claim
about the complete unified-operations goal. Never commit the compiled binary
or raw measurement output.
