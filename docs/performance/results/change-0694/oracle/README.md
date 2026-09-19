# Change 0694 shared-MCE oracle

This directory contains a standalone probe and a deterministic corpus driver
for the public `litchi_ooxml_common::mce::process_markup_compatibility` API.
It is evidence tooling only; it does not modify production sources and does
not make a performance claim.

Build the probe from the repository root after the baseline source checkout has
been frozen:

```sh
cargo build --release --locked \
  --manifest-path docs/performance/results/change-0694/oracle/Cargo.toml \
  --target-dir "$TARGET_DIR" -j 2
```

The binary is `mce-oracle-0694`. The package has its own empty workspace. Its
lockfile started from the neighboring change-0694 probe lockfile, then was
refreshed offline for the oracle package and direct path dependency. Baseline
and candidate builds use the same final lockfile; pass their absolute binary
paths to the driver.

The direct probe modes are:

```text
mce-oracle-0694 profiles
mce-oracle-0694 run <profile> <xml-path>
mce-oracle-0694 time <profile> <xml-path> <warmups> <samples>
```

`run` reads one XML file and prints one tab-separated record. An accepted input
prints the exact output byte length, SHA-256, ownership bit, and derived
`Report`; a refusal prints the typed Rust error's `Debug` representation. The
driver compares these records byte-for-byte and never canonicalizes or parses
the output XML.

The profiles are `baseline` (default OOXML capabilities and limits), `opaque`
(one explicit opaque extension name), `opaque-small` (that extension with
small input/output, depth, namespace, directive, and choice bounds),
`opaque-large` (the same extension with larger bounded values), and
`opaque-many` (4,096 explicit extension names). All profiles use the same
public processor entry point.

Run the bounded corpus differential with checked-in format fixtures:

```sh
python3 docs/performance/results/change-0694/oracle/corpus.py \
  "$BASELINE_ORACLE" "$CANDIDATE_ORACLE" "$REPO/test-data" \
  "$RUN_DIR" --timing
```

The default corpus selects four archives per format and three XML members per
archive, sorted by relative archive/member identity. It adds four mutations
per selected member (QName binding, duplicate attribute, unbound prefix, and
directive attributes) plus fixed cases for alternate content, opaque inactive
branches, an invalid descendant QName under an opaque parent, depth,
directive-token, and choice limits. This normally yields
roughly 150–250 cases and five profiles, while each member is capped at 128
KiB. The limits and budgets are command-line parameters when a smaller smoke
run is needed.

The output directory retains `corpus.json` with archive/member identities,
source and case hashes, and an aggregate corpus hash; `invocations.jsonl` with
one receipt per binary/profile/case invocation; `results.json` with exact
baseline/candidate mismatches and binary hashes; and, with `--timing`,
`timings.json` containing secondary wall-clock samples for the default,
small-extension, and large-extension profiles. Timing records are diagnostic
only and are not part of the equality decision.
