# Paired DOC paragraph replacement

This standalone harness compares one native variable-width body paragraph
replacement with one unified `File::set` call.  Each generated legacy DOC has
`N` plain body paragraphs named by their deterministic seed text.  The target
is the zero-based paragraph at `floor(N / 2)`, addressed through the unified
one-based path `/body/p[floor(N / 2) + 1]`.  The replacement is longer than the
seed, so the sample exercises the native CLX/piece-table rebuild and the
facade's corresponding body-text publication path.

The default matrix uses `N = 1, 10, 100, 1000` and nine ABBA rounds.  The
corpus is produced through the canonical `litchi_doc::Builder::build()` then
`Package::to_bytes()` publication boundary.  A fresh unified `File` is opened
from that corpus with source retention and `SourceMode::Auto` before every
unified timed sample, so retained edit state cannot warm later samples.  Every native sample opens a
`body_text::Snapshot` from the prepared source bytes inside the clock.  Every
unified sample then calls `File::set` against a parsed path; the DOC facade
opens the native package transaction and its variable-width body-text engine
inside that call.  File opening remains outside timing.  Path parsing, input
corpus construction, warm-up, post-operation output serialization, paragraph
readback, exact-byte comparison, and destruction are outside the timed
interval.  This is one paragraph operation, so the harness makes no claim
about batch or fanout scaling.

Run from the repository root with no concurrent build, test, or measurement
job:

```sh
BUILD=/tmp/litchi-unified-doc-paragraph-target
BIN="$BUILD/release/unified-doc-paragraph-baseline"
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR="$BUILD" cargo build --release \
  --manifest-path tools/perf-baseline/unified-doc-paragraph/Cargo.toml
LITCHI_SOURCE_COMMIT="$(git rev-parse HEAD)" \
LITCHI_RUSTC_VERSION="$(rustc --version)" \
LITCHI_TARGET="$(rustc -vV | sed -n 's/^host: //p')" \
LITCHI_SOURCE_DIGEST="$(git rev-parse HEAD^{tree})" \
LITCHI_MACHINE="taskset-cpu-0" \
LITCHI_BINARY_PATH="$BIN" \
LITCHI_BINARY_BYTES="$(stat -c '%s' "$BIN")" \
LITCHI_BINARY_SHA256="$(sha256sum "$BIN" | awk '{print $1}')" \
LITCHI_BUILD_COMMAND='CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=/tmp/litchi-unified-doc-paragraph-target cargo build --release --manifest-path tools/perf-baseline/unified-doc-paragraph/Cargo.toml' \
taskset -c 0 "$BIN" --rounds 9 > /tmp/unified-doc-paragraph.json
```

For a recorded run, retain the resolved lockfile digest alongside the source
manifest and use that same lockfile for both measurements.

The JSON receipt records the source and binary fingerprints, corpus byte
lengths and SHA-256 values, exact warm-up equivalence, paragraph text digests,
the timed boundaries, all raw samples, medians, and a failure matrix.  A
receipt with a corpus or output validation failure is not a successful timing
record.  Do not commit the benchmark binary or large raw captures.

## Diagnostic controls

Pass `--diagnostic` to add native-only controls after each paired ABBA round.
The policy control compares explicit default options with the facade's body
open/transaction policies. The retained-reader control runs a separate ABBA:
plain native edit versus the same edit with a freshly opened native Package
and Document kept alive outside the edit clock. It never preopens a body-text
Snapshot outside timing. Neither control contributes samples to the primary
route medians. Diagnostic runs can perturb allocator/cache state and are
investigative evidence, not replacements for ordinary acceptance runs.

Record 0189 contains repeated diagnostic receipts. Their source manifest
includes the external resolved lockfile; reconstruct it from the root lockfile
with the cached cc 1.6.0 resolution used for the offline build.
