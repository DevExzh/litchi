# Paired XLSB sheet-name rename

This standalone harness compares the native XLSB structural rename with one
unified `File::set` call. It generates one worksheet named `Sheet1`, fills it
with row-major numeric cells, and renames the first worksheet to `Renamed`.
The default corpus sizes are `1,10,100,1000` stored cells, both native eager
and native source owners, and nine ABBA rounds. Three rounds are the minimum.

The native eager leg edits a prepared `litchi_xlsb::Package`. The native source
leg starts from a prepared `SourceBackedWorkbook` and includes
`to_workbook().into_package()` inside every timed edit before calling
`Package::edit`, `set_sheet_name`, and `Commit::into_package`. Each unified leg
uses the matching prepared `File`: `SourceMode::Eager` for eager and
`SourceMode::Auto` for source, then times one `File::set("/sheet[1]", props)`
call. Opening, corpus generation, target/property preparation, warm-up,
serialization and equivalence checks are outside the clocks.

Run from the repository root with no concurrent build, test or measurement
job:

```sh
BUILD=/tmp/litchi-unified-xlsb-sheet-name-target
BIN="$BUILD/release/unified-xlsb-sheet-name-baseline"
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR="$BUILD" cargo build --release \
  --manifest-path tools/perf-baseline/unified-xlsb-sheet-name/Cargo.toml
LITCHI_SOURCE_COMMIT="$(git rev-parse HEAD)" \
LITCHI_RUSTC_VERSION="$(rustc --version)" \
LITCHI_TARGET="$(rustc -vV | sed -n 's/^host: //p')" \
LITCHI_BINARY_SHA256="$(sha256sum "$BIN" | awk '{print $1}')" \
LITCHI_BUILD_COMMAND='CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=/tmp/litchi-unified-xlsb-sheet-name-target cargo build --release --manifest-path tools/perf-baseline/unified-xlsb-sheet-name/Cargo.toml' \
taskset -c 0 "$BIN" --rounds 9 > /tmp/unified-xlsb-sheet-name.json
```

Override the matrix with:

```sh
unified-xlsb-sheet-name-baseline --counts 1,1000 --owners eager,source --rounds 3
```

The JSON records the generated corpus byte count and SHA-256, source commit,
compiler/target/build command, benchmark binary SHA-256, raw ABBA nanosecond
samples, upper-middle medians, ratios, native rename/commit accounting, and
the complete warm-up equivalence report. Exact
ZIP bytes are preferred. If they differ, every decoded ZIP member is compared
and the report names each missing or differing member with its byte length and
SHA-256; a mismatch exits nonzero. The result is scoped measurement evidence,
not a claim about the complete unified-operations goal. Never commit the
compiled binary or raw measurement output.

The native call uses the canonical complete-catalog selector API as of record
0183. Historical route names and JSON rename-call labels remain stable.
