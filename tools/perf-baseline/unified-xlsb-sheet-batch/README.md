# Paired XLSB direct sheet-path visibility batch benchmark

This standalone harness compares one native XLSB structural edit with one
unified `File::edit` transaction for hiding a fanout of worksheet tabs. Each
generated corpus begins with one visible, active `VisibleSentinel` tab and
then has `N` ordinary target worksheets. The native leg creates one
`Package::edit`, stages `N` native `set_sheet_state(..., State::Hidden)` calls
for catalog positions `1..=N`, and commits once. The unified leg prepares `N`
typed direct paths (`/sheet[2]` through `/sheet[N+1]`) and `N`
`visibility="hidden"` property sets, then submits all of them through one
`File::edit` closure. It deliberately does not use a query fanout. Keeping
the active sentinel first means every changed target is non-active, so this
workload exercises the homogeneous batch path without an active-view
relocation barrier.

The default matrix measures `N=1,10,100,1000,10000` for eager and source
owners, with nine ABBA rounds. Three rounds are the minimum. Opening owners,
parsing and typing every direct path, preparing properties, warm-up, output
serialization and equivalence checks are outside the timer. Source-native
samples promote `SourceBackedWorkbook` into a package inside the timed edit;
source-facade samples use the matching source owner and retain promotion in the
timed `File::edit` boundary. The JSON records the facade's retained commit
count, so a future active-view or publication barrier remains visible in the
evidence rather than being hidden by the one-call wrapper.

Run from the repository root with no concurrent build, test or measurement job:

```sh
BUILD=/tmp/litchi-unified-xlsb-sheet-batch-target
BIN="$BUILD/release/unified-xlsb-sheet-batch-baseline"
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR="$BUILD" cargo build --release --locked \
  --manifest-path tools/perf-baseline/unified-xlsb-sheet-batch/Cargo.toml
LITCHI_SOURCE_COMMIT="$(git rev-parse HEAD)" \
LITCHI_SOURCE_DIGEST="$(git diff HEAD --binary | sha256sum | awk '{print $1}')" \
LITCHI_RUSTC_VERSION="$(rustc --version)" \
LITCHI_TARGET="$(rustc -vV | sed -n 's/^host: //p')" \
LITCHI_MACHINE="$(hostname)" \
LITCHI_BINARY_SHA256="$(sha256sum "$BIN" | awk '{print $1}')" \
LITCHI_BUILD_COMMAND='CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=/tmp/litchi-unified-xlsb-sheet-batch-target cargo build --release --locked --manifest-path tools/perf-baseline/unified-xlsb-sheet-batch/Cargo.toml' \
taskset -c 0 "$BIN" --rounds 9 > /tmp/unified-xlsb-sheet-batch.json
```

Override the matrix with:

```sh
unified-xlsb-sheet-batch-baseline --counts 1,1000 --owners eager,source --rounds 3
```

The JSON records the generated corpus byte count and SHA-256, source commit,
compiler/target/build command, benchmark binary SHA-256, raw ABBA nanosecond
samples, upper-middle medians and ratios, native setter/commit accounting,
unified direct-operation/retained-commit accounting, and warm-up equivalence.
Exact ZIP bytes are preferred. If they differ, every decoded ZIP member is
compared and the report names each missing or differing member with its byte
length and SHA-256; a mismatch exits nonzero. The output also verifies every
target is hidden and the sentinel is visible after each route. This is scoped
measurement evidence, not a claim about the complete unified-operations goal.
Never commit the compiled binary or raw measurement output.
