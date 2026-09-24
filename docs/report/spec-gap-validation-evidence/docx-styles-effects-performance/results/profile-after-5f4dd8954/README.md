# DOCX `stylesWithEffects` matched after profile

This directory retains the bounded after profile for production source commit
`506e6f8e5fb94a14c390fc262b152c773e8990b1`, captured from clean checkout
`5f4dd8954acef1fea9e9d7de01c90b3f3f0c2721`. The checkout descends from the
current 52-lane correctness smoke from frozen harness HEAD
`0aa99593cb609487cf85230c361be23ca34d4367`. Its raw result retention is
committed at `421ff35ce710ab73777a8b260b5a4e26f3eef912`, and the independent
root verifier replay is retained at `c1003473d1aed4cf3019a8fcbfd95900f9ba31e5`.
The profile source bundle, `captured-source.bundle`, contains the after
checkout and requires the root-proof commit as its parent boundary.

The runner and an independent root replay both passed the 31 lane/scale rows,
three fresh processes, two warmups, and twenty measured samples per process:
1,860 measured samples and 186 warmups in total. The timing matrix used the
native fixture and generated 64 KiB and 1 MiB fixtures. The separate 52-lane
smoke supplied refusal, malformed-input, signed, absent-owner, inverse,
independence, and boundary coverage; no 8 MiB or near-limit timing was run.
The raw capture is 307 files totaling 12,383,248 bytes after adding the
independent `root-verification.json` receipt. It includes eight generated
fixture files and the empty per-process stderr sidecars. Every raw file's
size and SHA-256 is listed in `root-retention-verification.json`.

`verification.json` and `root-verification.json` are byte-identical. Their
SHA-256 is
`1a0c174ee77a270f1b54f9058cfa082ca767eb6785cc966db931cbbadbe571e6`.
The profile's generated inputs, source manifests, normalized Cargo metadata,
allocator phase definitions, and receipt file set were checked against the
before profile. The before profile is retained at
`../profile-clean-bf21536b3/`; its production source is
`8702fd4db8723acceb7deb51bcb40ff66604bf10` and its clean checkout is
`bf21536b35380005041917e7aee03a39804b2659`.

## Phase and memory interpretation

Preparation generated and hashed the XML, DOCX packages, replacement
resources, limits, expected member digests, and semantic inputs before each
operation clock. Capture and projection include their named package work.
Prepared mutation lanes keep setup outside the operation clock. The outer
`elapsed_ns` and phase clocks are separate from report writing. `publish_ns`
contains the public mutation/publication operation and contains nested
`apply_ns`; `apply_ns` includes the public styles-with-effects patch and put
path. `serialize_ns` covers package serialization after mutation. `reopen_ns`
and `inverse_reopen_ns` are separate readback phases. Nested phase values are
not additive with their containing phase.

Allocator fields report calls, requested/direct/reallocated/deallocated bytes,
live bytes, and peak live delta for the Rust counting allocator. They describe
the measured allocator window and are distinct from process RSS. The
`/usr/bin/time -v` sidecar records one maximum RSS marker for each fresh
process invocation, including its warmups and samples. Neither metric is a
managed process-memory cap.

The after profile's representative 1 MiB replacement medians were 155.671 ms
for main and 158.782 ms for glossary. For main, `apply_ns` was 90.037 ms,
`publish_ns` 93.733 ms, `serialize_ns` 3.708 ms, `reopen_ns` 29.203 ms, and
`snapshot_ns` 28.754 ms. For glossary those values were 92.066, 95.767,
3.700, 29.765, and 29.339 ms. The corresponding after allocator medians were
130,726,897 requested bytes and 5,767,866 peak-live-delta bytes for main, and
132,107,594 requested bytes and 5,790,767 peak-live-delta bytes for glossary.

## Matched before/after observations

`matched-comparison.json` contains all 31 lane/scale rows, all 60 measured
samples per row in each capture, p50/p95 values, allocator fields, phases, and
process-level RSS markers. Each ratio below is after p50 divided by before p50
for the same lane and scale. The values are bounded observations for these
fixtures and host runs; they do not establish a causal speedup, general
scaling law, production workload result, or all-workflow improvement.

| lane | before -> after elapsed p50 (ms) | ratio | apply p50 (ms) | publish p50 (ms) | peak live delta (bytes) |
| --- | ---: | ---: | ---: | ---: | ---: |
| replace main, 64 KiB | 21.513 -> 18.399 | 0.8553 | 13.118 -> 9.935 | 13.418 -> 10.241 | 799,352 -> 799,352 |
| replace main, 1 MiB | 186.614 -> 155.671 | 0.8342 | 120.668 -> 90.037 | 124.371 -> 93.733 | 5,864,443 -> 5,767,866 |
| replace glossary, 1 MiB | 189.680 -> 158.782 | 0.8371 | 122.329 -> 92.066 | 126.136 -> 95.767 | 5,865,059 -> 5,790,767 |
| capture synthetic main, 1 MiB | 31.361 -> 31.197 | 0.9948 | — | — | 4,637,484 -> 4,637,484 |
| projection main, 1 MiB | 29.864 -> 29.628 | 0.9921 | — | — | 2,841,345 -> 2,841,345 |
| noop main, native | 22.568 -> 23.208 | 1.0284 | 11.039 -> 11.055 | 11.070 -> 11.055 | 559,003 -> 559,003 |
| inverse replace main, native | 58.929 -> 49.498 | 0.8400 | 16.111 -> 15.157 | 16.597 -> 15.332 | 774,411 -> 774,411 |

The 1 MiB replacement rows retain the largest repeated phase cost in this
matrix: the after `apply_ns` and `publish_ns` medians remain much larger than
serialization. A narrow follow-up may inspect repeated traversal, copying,
and transient allocation in that public apply/put path. The phase is
composite and nested, so these receipts identify a measurement target rather
than a single function or a completed optimization decision.

The source changed between captures, while the scenarios, generated fixture
bytes, normalized dependency closure, toolchain, flags, allocator accounting,
phase boundaries, and semantic/opaque/inverse gates were matched. The
production source change and the changed source pin are recorded in the
receipts; the ratio table must therefore be read as a matched source-change
observation, not as an isolated causal attribution.

## Host and provenance limits

The host was Ubuntu 26.04.1 on AMD EPYC 9R45 with 32 logical CPUs,
129,447,068 KiB total memory, and affinity `0-31`. One-minute load was 0.95
before and 1.65 after this capture. Physical-core count and mount identity
were not captured. The host was shared and unisolated, so this evidence makes
no idle-host, dedicated-host, or uncontended timing claim. The retained
sanitized census records operator spot checks; it does not provide continuous
proof that no unrelated process overlapped the run. The capture was observed
with no competing compiler/profile process in the coordinator spot checks.

The recorded toolchain was Rust `rustc 1.95.0 (59807616e 2026-04-14)` and
Cargo `1.95.0 (f2d3ce0bd 2026-03-21)`, with target
`x86_64-unknown-linux-gnu`, the default linker, `CARGO_INCREMENTAL=0`,
`LC_ALL=C`, locked offline dependencies, and the recorded build environment.
The disposable profile binary was removed after the successful verification
sentinel. Its recorded SHA-256 is
`74e8a61744cf985d3003789c726e92be74e7d3baf20e6c20b39d5adbbe0a3a01`.
Build, source, command, fixture, host, metadata, allocator, and RSS receipts
remain here. This evidence makes no managed-memory-cap, native Word
acceptance, or later-production claim.

## Replay layout

The retained repository result directory cannot be replayed directly. The
runner's generated-fixture manifest contains absolute paths, and the
verifier's output-containment rule requires its evidence and replay output to
be under the recorded external results root. Restore the exact external path
`/var/tmp/litchi-docx-styles-effects-profile-results-5f4dd8954-20260912-a` and
copy the 307 raw paths listed in `root-retention-verification.json`, including
`generated-fixtures/`, without rewriting bytes or embedded absolute paths.
Write a new replay receipt inside that external directory; do not replace
`verification.json` or `root-verification.json`. The target may remain absent
after the successful cleanup sentinel.

Verify and fetch the retained source bundle, then use the captured HEAD:

```sh
git bundle verify /path/to/profile-after-5f4dd8954/captured-source.bundle
git fetch /path/to/profile-after-5f4dd8954/captured-source.bundle \
  HEAD:refs/heads/docx-styles-effects-profile-after-5f4dd8954
git worktree add --detach \
  /var/tmp/litchi-docx-styles-effects-profile-after-506e6f8e5-a \
  5f4dd8954acef1fea9e9d7de01c90b3f3f0c2721
```

The bundle SHA-256 is
`5b66dde8ddd50f7c57c11965fb0aa3afe27df629ece8bd34a3ed16a0705b9ad7`. It
verifies HEAD `5f4dd8954acef1fea9e9d7de01c90b3f3f0c2721` and requires parent
boundary `c1003473d1aed4cf3019a8fcbfd95900f9ba31e5`. Run the captured
checkout's `verify_profile.py` with that checkout as `--root`, its evidence
subtree as `--evidence`, the restored external directory as `--results`, and
replay output inside that same external directory. `root-retention-verification.json`
records the runner/root verification objects, raw-file hashes, source-bundle
hash, and comparison/report hashes. The target and external raw copy should
be removed only after a byte-identical replay and a durable Git audit.
