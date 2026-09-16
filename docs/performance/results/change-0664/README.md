# Change 0664 evidence packet

Marker-bearing PPTX and DOCX corpora with their marker-stripped controls, the
selectors over them, allocation metrics for the ordinary-save family, and the
two DOCX text-sink selectors change 0643 named as missing.

**No file under `crates/` was modified.** `git diff --stat 70d7768cc` touches
only `tools/`, `docs/` and this packet, so the library this baseline measures is
the base commit's library; there is no before/after pair to build and none was
built.

## Provenance

| | |
|---|---|
| Base commit | `70d7768cc` (`feat/office-format-completeness`) |
| Branch | `perf/0664-perf-harness-marker-bearing-corpora-and-save-allocations` |
| Host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| Toolchain | rustc 1.95.0, cargo 1.95.0 |
| CPU pin | `taskset -c 19` on every measured process, while seven other agents built and measured on the other cores |
| Binaries | staged outside every Cargo target directory; SHA-256 in [`timing/binary-sha256.txt`](timing/binary-sha256.txt) |
| Protocol | 20 warm-ups, 50 retained samples, three repeats (R1, R2, R3) per group, plus a dedicated A/A pair for the two heaviest groups |

`lto = true` makes release binaries non-reproducible byte for byte (change
0635), which is why the SHA-256 of each binary actually timed is recorded.

## Contents

| Path | What it is |
|---|---|
| [`scripts/derive_marker_shape.py`](scripts/derive_marker_shape.py) | the derivation: censuses the three real fixtures, emits the shape parameters, and `verify` re-derives them and fails if `marker_shape.rs` drifts |
| [`scripts/run_baseline.sh`](scripts/run_baseline.sh) | the capture driver: builds, stages, pins, runs the allocator legs first and the timing legs last |
| [`scripts/summarize.py`](scripts/summarize.py) | reduces the reports to the tables the record quotes |
| [`census/marker-shape.json`](census/marker-shape.json) | the derived shape: per part kind, the root declaration list, the `mc:Ignorable` value and the `mc:AlternateContent` template |
| [`census/pptx-real-deck-census.json`](census/pptx-real-deck-census.json) | every member of change 0649's deck, with its marker facts |
| [`census/docx-real-file-census.json`](census/docx-real-file-census.json) | every member of the real Word fixture |
| [`census/pptx-notes-fixture-census.json`](census/pptx-notes-fixture-census.json) | every member of the notes-bearing PPTX fixture |
| [`census/marker-census.json`](census/marker-census.json) | the harness's own per-member census of all four generated corpora (`--marker-evidence`) |
| [`timing/`](timing) | every retained report, one per group per repeat, plus the allocator reports |
| [`timing/summary.json`](timing/summary.json), [`summary.txt`](summary.txt) | the reduced tables |
| [`gates.txt`](gates.txt) | the tail of every gate |
| [`decision.json`](decision.json) | the decision record |
| [`cleanup.json`](cleanup.json) | what was removed |
| [`log-sections.md`](log-sections.md) | the four program-log paragraphs for the coordinator to merge |

## Replay

```sh
# the derivation, from the repository root
python3 docs/performance/results/change-0664/scripts/derive_marker_shape.py derive --out /tmp/census
python3 docs/performance/results/change-0664/scripts/derive_marker_shape.py verify

# the capture
bash docs/performance/results/change-0664/scripts/run_baseline.sh \
  <worktree> <staging-dir> <output-dir> 19
python3 docs/performance/results/change-0664/scripts/summarize.py <output-dir>
```

## The result, in one table

| | p50 (ms) | repeat spread | A/A floor |
|---|---:|---:|---:|
| `pptx_marker_ordinary_save_edit` (change 0649's phase) | **150.368** | 3.39% | **0.73%** |
| `pptx_marker_control_ordinary_save_edit` | 7.432 | 87.38% † | 86.56% † |
| `pptx_ordinary_save_edit` (change 0638's generated corpus) | 1.948 | 159.72% ‡ | — |

**77.21× the generated corpus. 20.23× a byte-identical control.** Change 0649
measured 77.3× (capture) and 67.7× (edit) on the real deck with a different
instrument, decomposing 5.57 (bytes) × 13.9 (per byte); this pair decomposes
3.82 × 20.23 = 77.3.

† and ‡ are single-repeat outliers on a shared host, not rounding: the control
edit's five p50s are 7.400, 7.409, 7.432, 7.677 and 13.866 ms, and the generated
edit's three are 1.928, 1.948 and 5.008 ms. The medians exclude them, both
within-run distributions are tight, and every series is published in
[`timing/summary.json`](timing/summary.json).

## Corpora

| Corpus | Members | Uncompressed | Marker-bearing | Share | `mc:AlternateContent` | Archive SHA-256 |
|---|---:|---:|---:|---:|---:|---|
| `pptx-marker-deck-marker` | 63 | 351,108 | 27 | 89.66% | 13 | `09d80c6ccb1159fe3abfe56db0aaf2c548f4bbd7147a9ae3223bb50c14b112d5` |
| `pptx-marker-deck-control` | 63 | 351,108 | 0 | 0.00% | 13 | `3a33dd743d91f05a7a207dd537342154816d8e16993e84552519dea04c486070` |
| `docx-marker-medium-marker` | 12 | 72,956 | 6 | 84.01% | 0 | `1762bd514f05cb944874bd318f2a9302786e99f879a93ea46a59e2ee1ec67711` |
| `docx-marker-medium-control` | 12 | 72,956 | 0 | 0.00% | 0 | `c61354b61d9b9b82c0baf472289e09f643d772c94ee76f335ffdfad9e09d38a4` |

The two corpora of a pair have the **same** uncompressed byte total and
different archive hashes, because the control substitutes an equal-length inert
URI for the namespace and deflate then sees different bytes.

Derivation fixtures (all ordinary tracked files):

| Fixture | SHA-256 | Members | Uncompressed | Marker-bearing | Share |
|---|---|---:|---:|---:|---:|
| `test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx` | `aebee5a724a9b5b8020c41fe789df9b3805265f20bf9a87dd28d8b0ed556deb0` | 103 | 796,725 | 43 | 93.03% |
| `test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/layout-in-cell-2.docx` | `4d273731897cb642d11062bf2b8ca8ec1b8f005843d51722172d1896e3a22cad` | 21 | 463,565 | 10 | 95.01% |
| `test-data/libreoffice-core/sd/qa/unit/data/pptx/tdf89064.pptx` | see `census/pptx-notes-fixture-census.json` | 42 | 68,699 | 17 | 57.66% |

## What is not claimed

No timing, allocation, physical-I/O, cold-cache or speedup claim, and no
claim-registry entry. The four publication rows (`save-to-path` and
`serialize-to-sink`) run at A/A floors of 13.13% to 240.56% — two `fsync`s on a
device eight agents share — and are reported and **not relied on**; the
allocation counts are the evidence for those phases. The DOCX sink refusal is a
corpus fact and not a claim about real Word files: both real DOCX fixtures
censused here are admitted. The full list is in the record's *Limitations*.
