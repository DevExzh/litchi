# 0003 S1 — pre-M0a baseline of the §5.3 routes

This directory holds the "before" capture for plan
[0003](../../../0003-m0a-migration-plan.md) step S1. It covers every route in
§5.3 (open, list, read-one, full-text, detection) on the harness's default
corpora at the pre-M0a commit. S10 compares the post-S9 build against it with
ABBA runs on the same binary, corpus and host.

This is a descriptive baseline, not an optimization claim. The host was shared
(see Noise), so read the spread columns before any p50.

## Identity

| Item | Value |
|---|---|
| Git revision | `da4c207478eb9192ce2f258ac41920ce55e01ab3` (branch `uo/m0a-s1-baseline`); nothing under `crates/` or `tools/` changed. The reports say `git_worktree_dirty: true` only because this results directory was untracked during the capture. |
| Binary | `/home/zhuhe/code/litchi-worktrees/uo-m0a-baseline-bin/litchi-perf-baseline-da4c207478` |
| Binary SHA-256 | `9c887ee55436adc1d98431449fa8c123e9616e73166e6f604e4bd81d12979c04` (264,629,168 bytes; also `binary_identity` in every report) |
| `tools/perf-baseline/Cargo.lock` | `db62bdd9174484613acae1f49f8187cf105d0b4976e2fea4de91df7372a893d5` |
| Toolchain | rustc 1.95.0 (59807616e 2026-04-14), cargo 1.95.0 |
| Host | AMD EPYC 9R45, nproc 32, Linux 7.0.0-1013-aws, 123 GiB RAM; every child pinned to CPU 20 |
| Corpus identity | `414fc710132c5fdf76061b7a0922c99ef175fa58862a9c00bdee925b3db8ba7d`: SHA-256 of the sorted (selector, corpus, archive SHA-256) set (`analysis.json`). Every row carries its own `corpus_archive_sha256`. Each group's schema-2 catalog/content-set hash is in `analysis.json.corpus_catalogs` and is identical across repeats. |
| OLE2 fixtures | the five checked-in files that 0638 used; SHA-256 values are in `environment.json` |
| Capture window | 2026-10-04 18:08:52 to 19:14:18 UTC (R1–R3); R4–R5 followed directly after |

## Files

- `capture.py`: the driver. S10 reuses it unchanged with `--binary` pointing at each leg and a separate `--out`.
- `analyze.py`: writes `analysis.json` and `summary.csv` from `raw/` and `children.jsonl`.
- `environment.json`: build recipe, binary/lock/fixture hashes and host.
- `children.jsonl`: one record per child process, with argv, start/end UTC, exit code, 1-minute load average before and after, and GNU `time` peak RSS and context switches.
- `raw/R<n>-<group>.json.gz`: the harness reports, with every retained sample, `sample_order`, per-row corpus identity, the harness `environment` and `binary_identity`. They are gzipped (`gzip -9 -n`); `analyze.py` reads them directly.
- `raw/R<n>-<group>.corpus-manifest.json`: the schema-2 corpus catalog for each child. The detection group has none (see Problems).
- `analysis.json`, `summary.csv`: per-row pooled statistics.

## Commands

Build (offline, locked, the 0907 recipe):

```sh
cd /home/zhuhe/code/litchi-worktrees/uo-m0a-s1-baseline
env CARGO_BUILD_JOBS=6 CARGO_INCREMENTAL=0 CARGO_NET_OFFLINE=true \
  CARGO_PROFILE_RELEASE_CODEGEN_UNITS=1 CARGO_PROFILE_RELEASE_DEBUG=1 \
  CARGO_PROFILE_RELEASE_LTO=fat CARGO_PROFILE_RELEASE_OPT_LEVEL=3 \
  CARGO_TARGET_DIR=$PWD/target LC_ALL=C TZ=UTC \
  cargo build --manifest-path tools/perf-baseline/Cargo.toml --locked --offline \
  --release --bin litchi-perf-baseline
cp target/release/litchi-perf-baseline \
  ../uo-m0a-baseline-bin/litchi-perf-baseline-da4c207478
```

Capture (R1–R3 for all groups, then R4–R5 for the groups that had noisy rows):

```sh
B=/home/zhuhe/code/litchi-worktrees/uo-m0a-baseline-bin/litchi-perf-baseline-da4c207478
D=docs/unified-ops/results/0003/before
python3 $D/capture.py --binary $B --out $D --repeats 3 --cpu 20 --label before
python3 $D/capture.py --binary $B --out $D --first-repeat 4 --repeats 2 --cpu 20 \
  --label before-noisy-repeat \
  --groups pptx_file,xlsx,docx_read_one,odt_file,rtf,ole2_facade_large,odp_file,ods_file
gzip -9 -n $D/raw/R*-*[a-z_].json && gunzip $D/raw/*.corpus-manifest.json.gz
python3 $D/analyze.py $D
```

Each child is run as
`taskset -c 20 /usr/bin/time -v -o <t> <binary> --case <group selectors> --warmup W --samples N --workers 1 [--filesystem-cache warm] [--ole2-file F ...] --json <out> [--corpus-manifest <m>]`.
`pptx_file` and `docx_file` use W=5 and N=30, because the harness spawns a
fresh child per sample there and each sample costs about 1 s. Every other
group uses W=10 and N=100. The exact argv of each child is in
`children.jsonl`. Groups run round-robin (R1 runs every group, then R2, and
so on), so drift on the host spreads over all groups.

## Route mapping (plan §5.3 → real selectors at da4c207478)

The plan's line references (`lib.rs` :1729–1770, :2041–2046, :2084–2096,
:2215–2216, :2283–2286, :2324–2327, :2349–2367) are correct at this commit.
The wildcards resolve as follows:

| Plan name | Selectors measured |
|---|---|
| `docx_file_source_open`, `…_eager_open` | `docx_file_{eager,source}_open` |
| `docx_file_*_list_paragraphs`, `…_paragraph_count[_lifecycle]` | `docx_file_{eager,source}_{list_paragraphs,paragraph_count}`, `docx_file_{eager,source}_open_paragraph_count_lifecycle` (the lifecycle name has an `open_` infix) |
| `docx_file_*_full_text[_lifecycle]` | `docx_file_{eager,source}_full_text`, `docx_file_{eager,source}_open_full_text_lifecycle` (`open_` infix) |
| `pptx_file_{eager,source}_open`, `…_list_slides`, `…_slide_count[_lifecycle]`, `…_selected_slide[_lifecycle]` | the 12 `pptx_file_*` selectors; the lifecycle names are `pptx_file_*_open_{slide_count,selected_slide}_lifecycle` |
| `xlsx_file_open[_lifecycle]`, `xlsx_bytes_open[_lifecycle]`, `xlsx_file_selected_cell` | as named |
| `doc_facade_*`, `ppt_facade_*` (opt-in) | as named, with `--ole2-file` and the 0638 fixtures |
| `rtf_file_open[_lifecycle]` | as named; default `plain` variant, all three semantic shapes |
| `odt_file_{eager,source}_open`, `odt_file_*_open_full_text_lifecycle` | as named |
| `odp_file_*_open`, `odp_file_*_selected_slide` | `odp_file_{eager,source}_{open,selected_slide}` |
| `ods_file_*_open`, `ods_file_*_selected_{cell,media}` | `ods_file_{eager,source}_{open,selected_cell,selected_media}` |
| `detect_{odt,ods,odp}_{bytes,reader,polyglot,catalog_alias}` | all 12, each on tiny/medium/large |
| **`docx_source_selected_paragraph` (bin)** | **This is not a timing route.** `src/bin/docx_source_selected_paragraph.rs` says it "deliberately does not measure elapsed time"; it reports ReadAt range evidence only. The docx read-one leg therefore uses the timed selectors `docx_semantic_one_paragraph` and `docx_semantic_one_paragraph_text` (leaf `litchi_docx`, tiny/medium/large) and `docx_producer_source_selected_paragraph` (source-backed, producer-shaped corpus). |

## Results

All times are µs. "n (repeats)" is the number of pooled retained samples
across repeats. The p50 CI is a 95% bootstrap interval that resamples within
each repeat, so it covers sampling noise only, not run-to-run drift. The
spread is (max − min) / median of the per-repeat p50s, and it is the noise
figure that matters for S10. `*_file_*` rows from fresh-child-per-sample
selectors are in the `warm` filesystem state.

| Scenario | Selector | Corpus | n (repeats) | p50 µs [95% CI] | p95 µs | spread R1–R3 | spread all |
|---|---|---|---|---|---|---|---|
| open | `doc_facade_file_open` | FloatingPictures.doc | 500 (5) | 186 [184–187] | 253 | 26.3% | 37.6% |
| open | `doc_facade_file_open` | duplicate-style-names.doc | 300 (3) | 11.55 [11.52–11.58] | 12.22 | 1.0% | 1.0% |
| open | `doc_facade_file_open` | documentProperties.doc | 300 (3) | 15.61 [15.58–15.66] | 16.53 | 1.8% | 1.8% |
| open | `docx_file_eager_open` | docx-source-backed-media | 90 (3) | 11,498 [11,436–11,569] | 13,045 | 0.1% | 0.1% |
| open | `docx_file_source_open` | docx-source-backed-media | 90 (3) | 316 [314–318] | 336 | 0.6% | 0.6% |
| open | `odp_file_eager_open` | odp-media-textbox-publication | 500 (5) | 174 [173–175] | 188 | 2.3% | 3.6% |
| open | `odp_file_source_open` | odp-media-textbox-publication | 500 (5) | 106 [105–106] | 121 | 9.8% | 12.1% |
| open | `ods_file_eager_open` | ods-media-publication | 500 (5) | 2,586 [2,581–2,591] | 2,703 | 5.3% | 5.2% |
| open | `ods_file_source_open` | ods-media-publication | 500 (5) | 1,998 [1,993–2,002] | 2,149 | 7.4% | 7.9% |
| open | `odt_file_eager_open` | odt-source-backed-repeated-text-media | 500 (5) | 2,056 [2,050–2,064] | 2,172 | 7.0% | 7.0% |
| open | `odt_file_source_open` | odt-source-backed-repeated-text-media | 500 (5) | 657 [656–658] | 671 | 3.4% | 3.4% |
| open | `ppt_facade_file_open` | SampleShow.ppt | 500 (5) | 30.43 [30.38–30.48] | 31.48 | 2.2% | 2.4% |
| open | `ppt_facade_file_open` | ppt_with_png.ppt | 300 (3) | 28.18 [28.09–28.22] | 30.04 | 0.4% | 0.4% |
| open | `pptx_file_eager_open` | pptx-source-backed-media | 150 (5) | 8,182 [8,123–8,227] | 8,879 | 4.4% | 7.3% |
| open | `pptx_file_source_open` | pptx-source-backed-media | 150 (5) | 2,835 [2,827–2,845] | 3,031 | 1.5% | 1.5% |
| open | `rtf_file_open` | rtf-semantic-plain-large | 500 (5) | 2,008 [1,999–2,016] | 2,156 | 4.4% | 8.0% |
| open | `rtf_file_open` | rtf-semantic-plain-medium | 500 (5) | 52.17 [52.00–52.28] | 57.77 | 1.1% | 3.2% |
| open | `rtf_file_open` | rtf-semantic-plain-tiny | 500 (5) | 7.39 [7.38–7.40] | 7.77 | 1.5% | 3.4% |
| open | `rtf_file_open_lifecycle` | rtf-semantic-plain-large | 500 (5) | 2,014 [2,010–2,017] | 2,109 | 5.2% | 5.2% |
| open | `rtf_file_open_lifecycle` | rtf-semantic-plain-medium | 500 (5) | 53.43 [53.29–53.56] | 59.62 | 0.4% | 3.4% |
| open | `rtf_file_open_lifecycle` | rtf-semantic-plain-tiny | 500 (5) | 7.66 [7.65–7.67] | 7.85 | 2.0% | 2.0% |
| open | `xlsx_bytes_open` | xlsx-cell-values-medium | 500 (5) | 62.18 [61.58–63.53] | 91.52 | 21.4% | 34.1% |
| open | `xlsx_bytes_open_lifecycle` | xlsx-cell-values-medium | 500 (5) | 1,805 [1,804–1,806] | 1,834 | 1.0% | 1.0% |
| open | `xlsx_file_open` | xlsx-cell-values-medium | 500 (5) | 323 [321–327] | 372 | 8.7% | 19.7% |
| open | `xlsx_file_open_lifecycle` | xlsx-cell-values-medium | 500 (5) | 2,781 [2,775–2,786] | 2,898 | 1.7% | 4.1% |
| list | `docx_file_eager_list_paragraphs` | docx-source-backed-media | 90 (3) | 146 [145–147] | 156 | 2.9% | 2.9% |
| list | `docx_file_eager_open_paragraph_count_lifecycle` | docx-source-backed-media | 90 (3) | 11,683 [11,603–11,762] | 13,635 | 1.8% | 1.8% |
| list | `docx_file_eager_paragraph_count` | docx-source-backed-media | 90 (3) | 99.18 [98.88–99.68] | 110 | 1.0% | 1.0% |
| list | `docx_file_source_list_paragraphs` | docx-source-backed-media | 90 (3) | 236 [234–239] | 279 | 3.4% | 3.4% |
| list | `docx_file_source_open_paragraph_count_lifecycle` | docx-source-backed-media | 90 (3) | 498 [495–500] | 520 | 1.5% | 1.5% |
| list | `docx_file_source_paragraph_count` | docx-source-backed-media | 90 (3) | 174 [173–176] | 195 | 1.4% | 1.4% |
| list | `pptx_file_eager_list_slides` | pptx-source-backed-media | 150 (5) | 5,946 [5,932–5,960] | 6,104 | 1.3% | 1.3% |
| list | `pptx_file_eager_open_slide_count_lifecycle` | pptx-source-backed-media | 150 (5) | 8,147 [8,103–8,178] | 9,078 | 3.8% | 5.4% |
| list | `pptx_file_eager_slide_count` | pptx-source-backed-media | 150 (5) | 3.38 [3.30–3.84] | 4.99 | 20.5% | 22.0% |
| list | `pptx_file_source_list_slides` | pptx-source-backed-media | 150 (5) | 6,169 [6,154–6,179] | 6,278 | 0.3% | 0.7% |
| list | `pptx_file_source_open_slide_count_lifecycle` | pptx-source-backed-media | 150 (5) | 2,894 [2,884–2,905] | 4,286 | 3.5% | 3.5% |
| list | `pptx_file_source_slide_count` | pptx-source-backed-media | 150 (5) | 2.04 [1.93–2.86] | 3.67 | 44.4% | 55.4% |
| read-one | `doc_facade_file_one_paragraph` | FloatingPictures.doc | 500 (5) | 494 [492–497] | 550 | 6.9% | 7.4% |
| read-one | `doc_facade_file_one_paragraph` | duplicate-style-names.doc | 300 (3) | 11.41 [11.38–11.44] | 11.87 | 3.0% | 3.0% |
| read-one | `doc_facade_file_one_paragraph` | documentProperties.doc | 300 (3) | 17.15 [17.13–17.18] | 18.06 | 2.1% | 2.1% |
| read-one | `docx_producer_source_selected_paragraph` | docx-producer-medium | 500 (5) | 106 [106–106] | 113 | 4.6% | 4.6% |
| read-one | `docx_semantic_one_paragraph` | docx-semantic-large | 500 (5) | 3,243 [3,242–3,245] | 3,335 | 0.8% | 1.1% |
| read-one | `docx_semantic_one_paragraph` | docx-semantic-medium | 500 (5) | 66.43 [66.39–66.51] | 72.82 | 1.3% | 2.0% |
| read-one | `docx_semantic_one_paragraph` | docx-semantic-tiny | 500 (5) | 9.71 [9.69–9.73] | 9.97 | 0.5% | 1.7% |
| read-one | `docx_semantic_one_paragraph_text` | docx-semantic-large | 500 (5) | 1.00 [0.99–1.02] | 2.00 | 64.6% | 65.2% |
| read-one | `docx_semantic_one_paragraph_text` | docx-semantic-medium | 500 (5) | 0.59 [0.59–0.60] | 0.69 | 5.0% | 5.8% |
| read-one | `docx_semantic_one_paragraph_text` | docx-semantic-tiny | 500 (5) | 0.57 [0.56–0.57] | 0.65 | 3.5% | 3.6% |
| read-one | `odp_file_eager_selected_slide` | odp-media-textbox-publication | 500 (5) | 77.05 [76.78–77.27] | 84.21 | 5.8% | 5.9% |
| read-one | `odp_file_source_selected_slide` | odp-media-textbox-publication | 500 (5) | 72.14 [71.91–72.38] | 79.41 | 6.5% | 9.0% |
| read-one | `ods_file_eager_selected_cell` | ods-media-publication | 500 (5) | 0.56 [0.56–0.58] | 0.82 | 25.5% | 25.5% |
| read-one | `ods_file_eager_selected_media` | ods-media-publication | 500 (5) | 275 [275–276] | 298 | 3.2% | 6.9% |
| read-one | `ods_file_source_selected_cell` | ods-media-publication | 500 (5) | 1.38 [1.35–1.39] | 1.68 | 15.8% | 15.8% |
| read-one | `ods_file_source_selected_media` | ods-media-publication | 500 (5) | 430 [428–430] | 444 | 5.3% | 5.5% |
| read-one | `ppt_facade_file_one_slide_text` | SampleShow.ppt | 500 (5) | 36.14 [36.11–36.23] | 37.78 | 0.8% | 3.0% |
| read-one | `ppt_facade_file_one_slide_text` | ppt_with_png.ppt | 300 (3) | 34.45 [34.40–34.48] | 36.13 | 1.2% | 1.2% |
| read-one | `pptx_file_eager_open_selected_slide_lifecycle` | pptx-source-backed-media | 150 (5) | 8,217 [8,153–8,308] | 9,566 | 4.8% | 6.2% |
| read-one | `pptx_file_eager_selected_slide` | pptx-source-backed-media | 150 (5) | 78.19 [77.53–79.02] | 88.25 | 2.2% | 7.0% |
| read-one | `pptx_file_source_open_selected_slide_lifecycle` | pptx-source-backed-media | 150 (5) | 2,976 [2,963–2,987] | 4,082 | 8.5% | 8.5% |
| read-one | `pptx_file_source_selected_slide` | pptx-source-backed-media | 150 (5) | 74.27 [73.64–74.87] | 87.65 | 1.7% | 5.9% |
| read-one | `xlsx_file_selected_cell` | xlsx-cell-values-medium | 500 (5) | 2,984 [2,981–2,986] | 3,092 | 0.2% | 0.9% |
| full-text | `doc_facade_file_full_text` | FloatingPictures.doc | 500 (5) | 278 [277–278] | 288 | 2.7% | 31.1% |
| full-text | `doc_facade_file_full_text` | duplicate-style-names.doc | 300 (3) | 11.47 [11.44–11.50] | 11.89 | 2.0% | 2.0% |
| full-text | `doc_facade_file_full_text` | documentProperties.doc | 300 (3) | 16.43 [16.38–16.48] | 17.24 | 2.6% | 2.6% |
| full-text | `docx_file_eager_full_text` | docx-source-backed-media | 90 (3) | 59.18 [58.99–59.33] | 67.69 | 1.3% | 1.3% |
| full-text | `docx_file_eager_open_full_text_lifecycle` | docx-source-backed-media | 90 (3) | 11,554 [11,428–11,674] | 12,278 | 3.0% | 3.0% |
| full-text | `docx_file_source_full_text` | docx-source-backed-media | 90 (3) | 146 [146–148] | 161 | 0.7% | 0.7% |
| full-text | `docx_file_source_open_full_text_lifecycle` | docx-source-backed-media | 90 (3) | 462 [459–463] | 492 | 1.6% | 1.6% |
| full-text | `odt_file_eager_open_full_text_lifecycle` | odt-source-backed-repeated-text-media | 500 (5) | 3,446 [3,438–3,454] | 3,699 | 7.2% | 7.2% |
| full-text | `odt_file_source_open_full_text_lifecycle` | odt-source-backed-repeated-text-media | 500 (5) | 1,805 [1,804–1,805] | 1,829 | 1.3% | 2.0% |
| full-text | `ppt_facade_file_full_text` | SampleShow.ppt | 500 (5) | 35.66 [35.62–35.73] | 36.84 | 1.3% | 1.6% |
| full-text | `ppt_facade_file_full_text` | ppt_with_png.ppt | 300 (3) | 32.06 [31.98–32.11] | 33.76 | 1.9% | 1.9% |

Detection: 300 samples per row (3 repeats); p50 in µs for large / medium / tiny
corpora, and the largest per-repeat p50 spread of the three corpora.

| Selector | p50 large / medium / tiny | max spread |
|---|---|---|
| `detect_odp_bytes` | 1.53 / 1.53 / 1.49 | 3.3% |
| `detect_odp_catalog_alias` | 12.64 / 12.51 / 12.58 | 4.5% |
| `detect_odp_polyglot` | 14.79 / 14.59 / 14.66 | 3.4% |
| `detect_odp_reader` | 6.67 / 6.62 / 6.58 | 2.6% |
| `detect_ods_bytes` | 1.11 / 1.09 / 1.08 | 1.9% |
| `detect_ods_catalog_alias` | 10.98 / 10.88 / 10.64 | 3.4% |
| `detect_ods_polyglot` | 12.85 / 12.78 / 12.63 | 2.3% |
| `detect_ods_reader` | 7.88 / 5.24 / 5.08 | 2.2% |
| `detect_odt_bytes` | 1.49 / 1.45 / 1.44 | 2.7% |
| `detect_odt_catalog_alias` | 12.41 / 12.37 / 12.36 | 3.0% |
| `detect_odt_polyglot` | 14.55 / 14.35 / 14.55 | 2.3% |
| `detect_odt_reader` | 7.12 / 6.53 / 6.44 | 2.2% |

### Peak RSS

The default (non-allocator) harness reports no per-route `process_metrics`
for these selectors. Whole-child peak RSS from GNU `time` is
**199.4–199.7 MiB for every group and repeat**, whatever the route, so it
measures the harness's fixed footprint and cannot tell routes apart. It is
recorded in `children.jsonl` and `summary.csv` only as a gross-regression
tripwire: S10 should flag a group whose whole-child peak moves by more than
about 1%. Per-route allocation needs the `litchi-perf-baseline-alloc` binary
(feature `allocator-metrics`). It was not built here because of the host disk
limit set during this step. If S10 needs per-route memory, build that binary
at both revisions.

## Noise

The host was shared with other agents' builds. The 1-minute load average
ranged from **8.3 to 135.2** on 32 CPUs. It peaked during R1's tail and R2's
`pptx_file` child (101 → 135); R2's later groups ran near 20, R3 near 41–44,
R4 at 40–66 and R5 at 22–49. Per-child values are in `children.jsonl`.
Involuntary context switches in the pinned child roughly doubled under the
peak (e.g. `pptx_file` 10.7k in R1, 20.4k in R2).

Groups with any row whose R1–R3 spread exceeded 5% were run twice more
(R4, R5): `pptx_file`, `xlsx`, `docx_read_one`, `odt_file`, `rtf`,
`ole2_facade_large`, `odp_file`, `ods_file`. Classified by the spread over
all repeats:

- **≤ 3% (usable against the 5% gate as is):** all detection rows except four
  between 3.0% and 4.5% (`detect_odp_catalog_alias` 4.5%, `detect_ods_catalog_alias` 3.4%, `detect_odp_polyglot` 3.4%, `detect_odp_bytes` 3.3%); every `docx_file_*` row except `docx_file_source_list_paragraphs` (3.4%) and `docx_file_eager_open_full_text_lifecycle` (3.0%);
  `pptx_file_source_{open,list_slides}`, `pptx_file_eager_list_slides`;
  `docx_semantic_one_paragraph`; `xlsx_bytes_open_lifecycle`,
  `xlsx_file_selected_cell`; `odt_file_source_open_full_text_lifecycle`;
  the small-fixture DOC/PPT facade rows.
- **3–10% (the 5% gate is at or below the noise):** `pptx_file_eager_open`
  (7.3%), the pptx lifecycles (5.4–8.5%), `pptx_file_*_selected_slide`
  (5.9–7.0%), `odt_file_eager_open` (7.0%) and its full-text lifecycle (7.2%),
  `ods_file_*_open` (5.2–7.9%), `ods_file_*_selected_media` (5.5–6.9%),
  `odp_file_*_selected_slide` (5.9–9.0%), `rtf_file_open*` large (5.2–8.0%),
  `doc_facade_file_one_paragraph` on FloatingPictures.doc (7.4%),
  `docx_producer_source_selected_paragraph` (4.6%), `xlsx_file_open_lifecycle`
  (4.1%), `docx_semantic_one_paragraph_text` medium (5.8%), and a tail of
  rows at 3.0–3.6% (`odt_file_source_open`, `odp_file_eager_open`,
  `pptx_file_source_open_slide_count_lifecycle`, `rtf_file_open*`
  medium/tiny, `ppt_facade_*` on SampleShow.ppt); `summary.csv` has every value.
- **> 10% (not gateable at 5% on this host):**
  - sub-5 µs reads where timer granularity and scheduling dominate:
    `pptx_file_{eager,source}_slide_count` (22%, 55%),
    `ods_file_{eager,source}_selected_cell` (25%, 16%),
    `docx_semantic_one_paragraph_text` large (65%);
  - `xlsx_bytes_open` (34%; R1 73.5 and R4 80.9 against about 60.3 in the other repeats) and
    `xlsx_file_open` (20%), both rising with load;
  - `odp_file_source_open` (12%);
  - **bimodal**: `doc_facade_file_open` on FloatingPictures.doc has run-level
    modes of about 180 µs (R1, R4, R5) and 246 µs (R2, R3), a 38% spread.
    `doc_facade_file_full_text` on the same file is about 277 µs in four
    repeats and 195 µs in R4. Within a run the mode is stable (tight CI), so
    the mode is set per process, not per sample. Do not compare single runs of
    these rows.

**Realistic S10 threshold.** With this noise, a flat 5% gate on single runs
would produce false positives on about a third of the rows. Recommended for
S10:
1. Run ABBA with at least two A and two B children per group, on the same CPU
   pin, and compare the median of per-child p50s, not pooled samples.
2. Per row, the regression trigger is `max(5%, spread_all from this table)`.
   For rows above 10%, also require that every B child's p50 is above every A
   child's p50 before a regression is called. Sub-5 µs rows get an absolute
   floor of 0.5 µs.
3. Record the load average for both legs (the driver does this). Re-run any
   ABBA block whose 1-minute load exceeds about 32, the CPU count.

## Problems and deviations

- `docx_source_selected_paragraph` is an untimed binary (see the mapping
  table), so it has no timing baseline. The substitutes are documented above.
- `--ole2-file` accepts **one file per format per child** ("accepts at most
  one file per format; … is the second"). The five 0638 fixtures are split into
  three children: `ole2_facade_small` (documentProperties.doc +
  ppt_with_png.ppt), `ole2_facade_large` (FloatingPictures.doc +
  SampleShow.ppt) and `ole2_facade_refusal` (duplicate-style-names.doc, DOC
  selectors only). 0638 records duplicate-style-names.doc as a frozen facade
  refusal; this report does not echo the outcome, so its rows time the refusal
  path.
- `--corpus-manifest` fails for the detection group: `ManifestError("content
  id odf-ooxml-zip:sha256:5f47…eb48 maps to conflicting V1 corpus objects")`,
  because the bytes and reader variants share an archive. Detection runs
  without the schema-2 manifest; per-row archive SHA-256 values are in its
  reports.
- Only the `warm` filesystem state is captured. The harness default also runs
  `cold-requested`, which is a request and not verified eviction (0907,
  README). The filesystem samples are on tmpfs (`filesystem_type: tmpfs`, no
  `--filesystem-root`).
- No route failed. 52 children ran and all exited 0.
