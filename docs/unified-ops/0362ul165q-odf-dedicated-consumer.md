# 0362ul165q — Isolated ODF raw-codec consumer

The ODF raw-codec qualification now uses
`tools/odf-raw-codec-consumer/Cargo.toml`. This isolated package enables all
ten ODF facade families without the unrelated OOXML dependencies that made
the historical `70d` benchmark package fail compilation. It references the
existing canonical probe; other performance consumers and their defaults
remain unchanged.

The preparer, runner, manifest, tests and gate lock registry migrate
together in leaf `fa1d7b6c47f45e32ee5c95a1f45448f9b0d7f821`.
The release profile retains LTO and abort-on-panic. Endpoint manifests
preserve each production source's compression dependency policy: the
historical source uses registry `flate2 1.1.10` and `zlib-rs 0.6.7`, while
the candidate uses its tracked vendored dependencies. Candidate vendor
files are not copied into the historical production source.

## Lock provenance and admission

The historical dedicated lock is generated offline from the historical
seed and pinned independently of the candidate lock. Preparation and runner
handoff verify its fixed identity and provenance, including source
commit/tree, seed, manifest, probe, metadata command and registry graph.
Candidate-lock substitution, arbitrary lock bytes, wrong endpoint policy,
missing vendor inputs and provenance drift are refused. These checks close
the independent review returns; there is no accepted baseline exception.

Root independently compares the seed and resolved lock. All 170 retained
registry versions, sources and checksums match the 204-entry seed. Thirty-four
registry packages and four optional dependency edges disappear with the
narrower feature set. One additional textual change removes an unnecessary
version qualifier from the same `spin 0.9.9` dependency; it is not a new
dependency. The candidate and historical locks remain distinct.

## Validation

Independent review passes the public preparation/runner flow and all 21
focused Python tests, including the gate lock registry. Four serial,
offline, locked release compiler checks then pass:

| Source | Feature selection | Result |
| --- | --- | --- |
| Historical `70d` | `raw-codec-original` | Pass |
| Historical `70d` | `raw-codec-original,raw-codec-allocator` | Pass |
| Candidate `fa1d7b6c47` | `raw-codec-current` | Pass |
| Candidate `fa1d7b6c47` | `raw-codec-current,raw-codec-allocator` | Pass |

Each command uses:

```sh
cargo check --release --offline --locked --jobs 1 \
  --manifest-path tools/odf-raw-codec-consumer/Cargo.toml \
  --example odf_raw_codec_qualification --no-default-features \
  --features <selection>
```

The clean historical compiler source is consumer-overlay commit
`28bdce66b652988a1e1be5358b82d5afcbe7929d` over
`70d64634c9d1d4328447d8724d3b076114a91fd0`. Root verifies that only the new
consumer manifest, its lock and the canonical probe differ. Production
crates, the root manifest and dependency implementation files are unchanged.

Root verifies all four logs and target traces, unchanged source and lock
identities, tool pins and absence of remaining builders. Targets remain
below 176 MB, within the 20 GB cap. These are compiler checks, including
compilation of the allocator variant; no probe, native timing or allocator
measurement has run. Actual final-C/P rebinding and qualification are still
required.

## Coordination

The DOCX repair has merged at `babf023e1a`; its worktree/ref and 850,676,076
apparent bytes of obsolete debug artifacts are retired. Logs, receipts and
the ignored root lock remain preserved. Cumulative cleanup reaches 281
worktrees and 443 refs. The post-merge census records 78 heads and 66
worktrees, including the active ODF and DOCX metadata qualification leaves;
all 70 excluded refs remain unchanged.

The historical 28 preparation groups are source-ready again. The new DOCX
metadata group is under independent review, making 29 required native
groups. No group has final-candidate acceptance yet. Final all-65, census,
corpus, fidelity and performance work remain required. Nothing is pushed.

## Evidence

Paths are relative to `/home/zhuhe/code/litchi-worktrees/.coord/`.

| Artifact | SHA-256 |
| --- | --- |
| `odf-raw-consumer-scope-independent-review-fa1d-20261026/review-receipt.json` | `37437c8e00ed590250b17edb4ebee8d5cc0cb3405f32d13605a0adfc4a91440e` |
| `odf-raw-original70d-dedicated-consumer-metadata-20261026-v1/baseline-lock-provenance.json` | `5d2f2d51b90fee928e34d14d04375f2b1341732a3d96aec89e50a56963039ac1` |
| `odf-raw-dedicated-lock-root-verification-20261026/receipt.json` | `a39dfaabc06bea2e97ed0bfc6748ca9cc2afd54da87f5ff0e716e0188f567737` |
| `odf-raw-dedicated-release-compile-fa1d-v2/result.json` | `b52b74153e5cb82e5557e9e0104454fa67010f20e339cffa46afd23a9563f219` |
| `odf-raw-dedicated-compile-root-verification-20261026/receipt.json` | `d2b79ab2df72f06e92c32cc850fffb0dd48454845ca0de230a1f261dec9e37f8` |
| `docx-metadata-repair-retirement-20261026/receipt.json` | `e7bb1f3eab31781694eea0653594272bfd7aa46713d3c83071407f0b7ad217f7` |
| `branch-worktree-165p-postmerge-20261026/receipt.json` | `29d9fb8c2f117f39c14375c3cd396d62984c9e78cac38fccb37ac246a68791be` |

Earlier failed compiler attempts remain preserved. Their results are not
relabelled or reused as passing qualification evidence.
