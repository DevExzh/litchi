# Paired RTF paragraph removal

This standalone harness compares one native RTF paragraph removal with one
unified `File::remove` path, and compares a native paragraph-removal batch with
one scoped unified query removal.  The native and unified owners are opened
once per case, warmed, and then reused as immutable inputs.  Input parsing,
selector parsing and corpus construction are outside the timed intervals.

The corpus is deterministic plain body RTF containing `N` paragraphs named
`paragraph 00000`, `paragraph 00001`, and so on.  The single-target case
removes the middle paragraph through `/body/p[N/2 + 1]`.  The query case
removes every paragraph through a `p` query scoped to `/body`; its native
counterpart calls `Edit::remove_body_paragraphs` with the original zero-based
positions in body order.  Each warm-up and timed result is reopened and
checked semantically, and the native and unified warm outputs must have
identical bytes.

The default matrix uses 10 and 1,000 source paragraphs, two untimed warm-ups,
and nine ABBA rounds.  The query case therefore measures one atomic removal
of 10 or 1,000 selected paragraphs, subject to the explicit native policy
profile of 10,000 operations and 16 MiB staged bytes.  Override the corpus
counts and iteration protocol with:

```sh
unified-rtf-remove-baseline --counts 10,1000 --warmup 2 --rounds 9
```

Run from the repository root with no concurrent build, test, or measurement
job:

```sh
BUILD=/tmp/litchi-unified-rtf-remove-target
BIN="$BUILD/release/unified-rtf-remove-baseline"
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR="$BUILD" cargo build --release \
  --manifest-path tools/perf-baseline/unified-rtf-remove/Cargo.toml
LITCHI_SOURCE_COMMIT="$(git rev-parse HEAD)" \
LITCHI_SOURCE_DIGEST="$(git rev-parse HEAD^{tree})" \
LITCHI_RUSTC_VERSION="$(rustc --version)" \
LITCHI_TARGET="$(rustc -vV | sed -n 's/^host: //p')" \
LITCHI_MACHINE="taskset-cpu-0" \
LITCHI_BINARY_PATH="$BIN" \
LITCHI_BINARY_BYTES="$(stat -c '%s' "$BIN")" \
LITCHI_BINARY_SHA256="$(sha256sum "$BIN" | awk '{print $1}')" \
LITCHI_BUILD_COMMAND='CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=/tmp/litchi-unified-rtf-remove-target cargo build --release --manifest-path tools/perf-baseline/unified-rtf-remove/Cargo.toml' \
taskset -c 0 "$BIN" --warmup 2 --rounds 9 > /tmp/unified-rtf-remove.json
```

The JSON receipt records the source and binary fingerprints, corpus and
output hashes, exact-byte equivalence, raw nanosecond samples, medians and
75th percentiles, and the fixed ABBA protocol.  A failed semantic or byte oracle
causes a nonzero exit status.  Timing summaries are descriptive; they do not
establish a speedup, allocation result, or production ranking.  Never commit
the compiled binary or a large raw capture.
