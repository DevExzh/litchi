# 0947 replacement-window coverage evidence

See [the report](../../0947-xlsx-replacement-window-coverage.md). These are
bounded review copies, not the complete runnable packet. The complete packet
retains fixtures, the isolated two-crate source export, probe binaries,
historical diagnostic outputs, and preparation versions outside Git.
Production source and publication behavior are unchanged.

The primary result is [coverage-model-v1.json](identity/coverage-model-v1.json),
computed from the current [proof output](results/proof-portable-final.tsv) and
[independent namespace-aware cell spans](review/changed-cell-span-review-v4.json).
It retains workbook auditing and the existing tiny windows. The model is
hypothetical coverage, not a semantic multi-window proof or measured speedup.

## External archive

- File: `/home/zhuhe/litchi-performance-0947-archive/litchi-performance-0947-v1.tar.gz`
- Bytes: **26,618,833**
- SHA-256: `885fe4aa7b81af78b2e08586c316780990735509532cb8133caaa310345a380a`
- Exact file inventory: [inventory-v1.json](inventory-v1.json), 150 files.
- Retained packet: `/home/zhuhe/litchi-performance-0947`.

Regenerate the archive from that unchanged packet into a fresh directory:

```sh
python3 -B docs/performance/results/change-0947/archive_packet.py \
  --packet /home/zhuhe/litchi-performance-0947 \
  --destination /tmp/0947-new-archive
```

Extract to an empty directory, then change into `litchi-performance-0947`.
Run with fresh output paths:

```sh
binaries/probe-v2 inputs/artifacts /tmp/0947-new-proof.tsv
cmp /tmp/0947-new-proof.tsv results/proof-portable-final.tsv
python3 -B analysis/analyze.py --artifacts inputs/artifacts \
  --manifest inputs/artifacts/manifest.json --probe /tmp/0947-new-proof.tsv \
  --output /tmp/0947-new-diagnostic.json
cmp /tmp/0947-new-diagnostic.json results/diagnostic-portable-final.json
python3 -B identity/coverage_model.py \
  --spans review/changed-cell-span-review-v4.json --proof /tmp/0947-new-proof.tsv \
  --output /tmp/0947-new-model.json
cmp /tmp/0947-new-model.json identity/coverage-model-v1.json
```

The full analyzer output remains external at
`results/diagnostic-portable-final.json`: **594,327 bytes**, SHA-256
`af40d5a8dc95e4cfa6a922431c7da98e746f966e5548626acb1cfc4815cedd93`.
The source/command/output bindings are in
[receipt-portable-v2.json](receipt-portable-v2.json).

To rebuild from the extracted source with the recorded toolchain and cached
locked dependencies:

```sh
CARGO_TARGET_DIR=/tmp/0947-new-target cargo +1.95.0 build \
  --manifest-path probe-v2/Cargo.toml --locked --offline
/tmp/0947-new-target/debug/litchi-proof-diagnostic inputs/artifacts \
  /tmp/0947-rebuilt-proof.tsv
cmp /tmp/0947-rebuilt-proof.tsv results/proof-portable-final.tsv
```

The retained probe binary is **7,916,440 bytes**, SHA-256
`5623a3784cb0dbe5e1b2e27372a49b36bcf771e6c276ca5f53471a69396d6105`.
Copied production crate files match source commit
`a6aca5c5f064c3fecc3edbf532c9637ee788e134`; the isolated workspace manifest
is diagnostic setup. Debug executable hashes can change with build paths;
reproduction compares the source/dependency bindings and diagnostic outputs.
No compiled executable or fixture ZIP is committed.

## Verification and limits

Final probe formatting, locked offline build, and warning-denied Clippy passed.
The root verified all archived file paths, sizes, and hashes after extraction,
then rebuilt at the moved location. The proof TSV, analyzer JSON, and coverage
model all reproduced byte for byte. The rebuilt probe refused an existing
output without changing it. See [root-replay-v1.json](root-replay-v1.json).
The owned extraction and rebuild target were removed after verification.

The independent Expat tool's copied-input replay matches after normalizing
only its embedded absolute identity-receipt path; its raw JSON is not
byte-identical after relocation. Five rejection controls pass: corrupt XML,
missing member, wrong identity hash, duplicate cell key, and output overwrite.
Its trusted identity receipt binds the XML; it does not validate arbitrary
fixture manifests. Root's [identity check](identity/verified-v1.json) separately
verifies the three source/output ZIPs and all twelve XML pairs against the
adopted 0945 candidate reports. Rerunning that identity check requires the
original repository/prior-packet inputs listed in its recorded command.

The analyzer's cell arrays match the independent spans exactly. Its additional
minimal-value-window estimates remain lexical diagnostics and are not used by
the conservative model. General producer namespaces, MCE, malformed inputs,
aggregate multi-window limits, and allocation behavior remain candidate
qualification work. The public probe reports proof kinds; fallback explanations
in the [source review](review/boundedsource-review.md) are source-derived.

The two-crate export omits a test-only OPC dependency, so a broader source
formatting attempt could not resolve that module. The final probe formatting
check passed; no full workspace formatting or Rust test pass is claimed.
Initial probe formatting was corrected before the final logged build.
Raw formatting diff logs remain unchanged in the external archive; the
committed copies omit those two files rather than alter their recorded bytes.
An initial identity check used an incomplete source path and failed before
writing its result; the corrected source path passed.

Preparation was exploratory. Not every intermediate analyzer source was
retained; historical JSON outputs are not claimed reproducible with the final
helper. An initial quadratic byte-diff attempt was stopped. Final v2 has
before/after source hashes and retained command logs. See
[preparation history](identity/preparation-history-v1.json). No native benchmark,
allocation, RSS, or publication performance measurement was performed.
