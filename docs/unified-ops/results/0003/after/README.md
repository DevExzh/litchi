# 0003 S10 — after-measurement of the §5.3 routes (ABBA)

This directory holds plan [0003](../../../0003-m0a-migration-plan.md) step
S10, the before/after comparison that gates the single M0a merge. Two
questions are answered for every §5.3 route:

1. **Facade comparison.** Did the route get slower? The old route runs on the
   pre-M0a binary and its `unified_` successor on the post-S9 binary.
2. **Overhead check.** Is the unified route within the goal's budget? The
   `unified_` route is compared with its `leaf_` twin, both on the post-S9
   binary.

Children alternate ABBA on one pinned CPU, and S1's regression rule is applied
to every row. The verdict and the explanations are in the record,
[`s10-measurement.md`](../../../s10-measurement.md).

No row here is an optimization claim. The host was shared, so read the
spread and trigger columns before any ratio.

## Summary

| Comparison | Rows | Result |
|---|---:|---|
| Facade (old selector → `unified_` successor) | 99 | 91 pass; 8 over their trigger, all explained (§"Follow-ups") |
| Kept predecessors (same selector on both binaries) | 47 | 47 pass (0.87–1.03) |
| Budget-only successors (timed scope differs) | 4 | no facade verdict |
| Overhead (`unified_` vs `leaf_`) | 85 | 25 open rows: 21 within the budget, 4 over (explained); 60 other rows: 37 within the review trigger, 23 review items (explained) |

The 8 facade rows over their trigger:

- **E1, first call in a fresh process.** Seven rows, all on routes that run
  one open or one query per fresh child process:
  - `xlsx_file_unified_open` (1.488) and `xlsx_file_unified_open_lifecycle`
    (1.058);
  - `pptx_file_unified_{eager,source}_slide_count` (2.56 and 2.77, +2.7 µs)
    and `pptx_file_unified_source_selected_slide` (1.091);
  - `docx_file_unified_eager_paragraph_count` (1.081) and
    `docx_file_unified_eager_full_text` (1.173). The full-text row also
    carries the S8a paragraph separator.

  Standalone probes built from the same sources with the same recipe show
  the same first-call time at both revisions. Callgrind shows the same work
  (+0.75% instructions on the XLSX open). The gap belongs to the harness
  executables, not to M0a's code.
- **E2, allocator state in the RTF sample loop.** `rtf_file_unified_open` on
  the 540 KB corpus (1.680). With malloc trimming off, or in a bare open loop,
  the unified, leaf and pre-M0a opens cost the same.

## Identity

| Item | Value |
|---|---|
| Before binary (leg A) | `/home/zhuhe/code/litchi-worktrees/uo-m0a-baseline-bin/litchi-perf-baseline-da4c207478`, SHA-256 `9c887ee55436adc1d98431449fa8c123e9616e73166e6f604e4bd81d12979c04` (264,629,168 bytes), checked against that directory's `SHA256SUMS`. S1 built it from `da4c207478`, which the 2026-10-05 history rewrite renamed `a8f45144c2` ([`history-rewrite-2026-10-05.md`](../../../history-rewrite-2026-10-05.md)). |
| After binary (legs U and L) | built from `d10dd4c747`: `uo/m0a` `b906691f3b` (S1–S9 and records 0011–0014 merged) plus the S10 harness fix (§"Harness fix"). SHA-256 `dbe581e730b159872165f871067f23e27ae4baedf7cdd6974f5a5939d6985761` (265,331,992 bytes). Not committed (AGENTS.md); rebuild with the command below. `c475a527df` adds only a `#[cfg(test)]` test. |
| Build recipe | S1's (`../before/environment.json`): `CARGO_BUILD_JOBS=6 CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_{CODEGEN_UNITS=1,DEBUG=1,LTO=fat,OPT_LEVEL=3}`, `--locked --offline --release --bin litchi-perf-baseline` |
| `tools/perf-baseline/Cargo.lock` | before `db62bdd9…72a893d5` (S1's); after `a9979b6e…6efe3f`. The only package difference is the added path package `litchi-ops`; every third-party version is identical. |
| Toolchain | rustc 1.95.0 (59807616e 2026-04-14), cargo 1.95.0, both legs |
| Host | AMD EPYC 9R45, 32 CPUs, 1 thread per core, Linux 7.0.0-1013-aws, 123 GiB RAM; every child pinned to CPU 20, as in S1. Shared with other agents. |
| Corpus | each binary generates its corpora. Every compared row has the same archive SHA-256 on both legs (`archive_sha256_match` in `analysis.json`). The five OLE2 fixtures match S1's hashes (`environment.json`). |
| Capture window | 2026-10-05 18:48:50 to 22:58:20 UTC: in-process groups and XLSX 18:48–19:35, PPTX 20:54–22:07, DOCX 22:07–22:58. The follow-ups ran 22:58–23:10. |

## Files

| File | What it is |
|---|---|
| `capture_s10.py` | The ABBA driver. It imports S1's `capture.py` unchanged (group map, child argv, samples, GNU `time` parser) and builds the U and L maps from `../s8b/route-map.json`. Added during the run: a retry of children that failed on the tmpfs quota, `--filesystem-root`, `--gzip-raw`. |
| `children.jsonl` | One record per child: argv, binary SHA-256, start and end, 1-minute load before and after, load-gate wait, attempt, exit code, GNU `time`. |
| `pruned-<group>.json.gz` | Every harness report, cut down by `prune_reports.py` to the fields the analysis reads: samples, corpus identity, output digests and observations, fresh-child peak RSS. Each file is under 160 KB. |
| `analyze_s10.py`, `analysis.json`, `summary.csv` | The S1 rule, the pairs, memory, output evidence, and per-row statistics and verdicts. It runs on this directory as committed. |
| `explanations.json` | The reviewed explanations of rows over their trigger. `analyze_s10.py` matches them by pattern. |
| `render_tables.py` | Renders the tables below. |
| `environment.json` | Binaries, probes, revisions, build recipe, host, fixtures, filesystem roots. |
| `mem_probe.rs`, `mem_probe_build.sh`, `mem_compare.py`, `memory-probe*.json*` | The allocation probe (R5) and its results (§"Memory"). |
| `cold_open_probe.rs`, `cold_runs.py`, `cold-open.json` | The fresh-process first-call probe and its results, with callgrind counts (§"Follow-up E1"). |
| `rtf_allocator_check.py`, `rtf_allocator_summary.py`, `rtf-allocator-check.json`, `rtf_open_probe.rs`, `rtf-open-probe.json` | The allocator follow-up for the RTF row (§"Follow-up E2"). |

The full raw reports are not committed. They are 8,826,880 bytes as
`s10-capture-raw.tar` in the session scratchpad, SHA-256
`70e44835ed96cc9d715c61bab18131d8931f5122064446a001ae03c0eb9676b6`, and hold
`capture/` and `rtf-check/`. That includes the superseded PPTX tmpfs
children (§"Problems").

## Commands

Build the after binary (S1's recipe, at `d10dd4c747`):

```sh
cd /home/zhuhe/code/litchi-worktrees/uo-m0a-s10-measure
env CARGO_BUILD_JOBS=6 CARGO_INCREMENTAL=0 CARGO_NET_OFFLINE=true \
  CARGO_PROFILE_RELEASE_CODEGEN_UNITS=1 CARGO_PROFILE_RELEASE_DEBUG=1 \
  CARGO_PROFILE_RELEASE_LTO=fat CARGO_PROFILE_RELEASE_OPT_LEVEL=3 \
  CARGO_TARGET_DIR=$PWD/target LC_ALL=C TZ=UTC \
  cargo build --manifest-path tools/perf-baseline/Cargo.toml --locked --offline \
  --release --bin litchi-perf-baseline
sha256sum target/release/litchi-perf-baseline   # dbe581e7…6985761
```

The capture as it ran. `OUT` is a scratch directory, `A` the binary above
and `B` the S1 binary:

```sh
D=docs/unified-ops/results/0003/after
B=/home/zhuhe/code/litchi-worktrees/uo-m0a-baseline-bin/litchi-perf-baseline-da4c207478
run() { python3 $D/capture_s10.py --before-binary $B --after-binary $A --out $OUT --cpu 20 "$@"; }
run --groups docx_read_one,odt_file,rtf,ole2_facade_small,ole2_facade_large,ole2_facade_refusal,detection --blocks 8
run --groups odp_file,ods_file --blocks 4
run --groups xlsx --blocks 3
# PPTX blocks 1-3 (tmpfs) were superseded (see Problems); the re-run:
run --groups pptx_file --blocks 2 --first-block 4 --filesystem-root $PWD/target/s10-fsroot --gzip-raw
run --groups docx_file --blocks 1 --filesystem-root $PWD/target/s10-fsroot --gzip-raw
```

Analysis, which reproduces from this directory as committed:

```sh
python3 $D/prune_reports.py $OUT/raw $D && cp $OUT/children.jsonl $D/
python3 $D/analyze_s10.py $D $D            # analysis.json, summary.csv
python3 $D/render_tables.py $D $D/children.jsonl
```

Follow-ups. Probes are built with `mem_probe_build.sh <repo> <before|after>
<work dir> <target dir> [source bin]`. For the pre-M0a builds `<repo>` is a
detached sparse worktree of `a8f45144c2` with its own target directory. The
timing probes add S1's `CARGO_PROFILE_RELEASE_*` variables.

```sh
python3 $D/rtf_allocator_check.py --before-binary $B --after-binary $A --out $OUT2 --cpu 20 --blocks 2
python3 $D/rtf_allocator_summary.py $OUT2
taskset -c 20 s10-rtf-open-probe large {bare,harness} 100 2   # also under GLIBC_TUNABLES
taskset -c 20 s10-rtf-open-probe large flat 200
python3 $D/cold_runs.py 20 30 out.json NAME=PROBE:FORMAT:MODE:FILE ...  # cold-open.json lists the variants
S10_ROUNDS={0,1,2} valgrind --tool=callgrind s10-cold-open-probe xlsx {facade,leaf} corpus.xlsx
taskset -c 20 s10-mem-probe 5 2 FORMAT:FILE ...   # S10_RETAIN_SOURCE=0 for retain_source(false)
python3 $D/mem_compare.py before.jsonl after.jsonl memory-probe.json after-retain-false.jsonl
```

The corpus files for the probes were copied from a harness run of the same
selector (`--filesystem-root`) and checked against the archive SHA-256 in the
reports. The probe fixtures are the S1 OLE2 files plus the files in
`memory-probe.json`.

## Protocol

- Legs. **A**: the pre-M0a binary on S1's selectors (S1's `capture.py`
  group map, imported unchanged). **U**: the post-S9 binary on the unified
  successors (`../s8b/route-map.json`) plus the 19 selectors S9 kept, which
  run in both binaries under one name. **L**: the post-S9 binary on the leaf
  twins.
- Order. Per S1 group, each block runs the children `A U L L U A`: `A U … U A`
  is the ABBA order of the facade comparison and `U L L U` that of the
  overhead check. A group's block is contiguous in time; groups run one
  after another.
- Blocks per group: 8 for the cheap in-process groups (DOCX read-one, ODT,
  RTF, the three OLE2 splits, detection), 4 for ODP and ODS, 3 for XLSX,
  2 for PPTX (blocks 4 and 5, see §"Problems") and 1 for DOCX, whose S1 noise
  is at most 3.4%. Every leg has at least two children per block.
- Children are S1's: `taskset -c 20`, GNU `time -v`, `--workers 1`, S1's
  warm-up and samples (5/30 for PPTX and DOCX, 10/100 otherwise),
  `--filesystem-cache warm` for the filesystem groups, a schema-2 corpus
  manifest except for detection.
- Load gate: before each child, while the 1-minute load average was above 8,
  the driver slept 30 s and re-checked, with a 2-hour cap on the total wait
  over the whole capture. Eleven children waited, 600 s in total; the cap was never reached.
- Statistic: per row and leg, the median of the per-child p50s (S1's p50 of
  each child's retained samples).
- Rule (S1): trigger = max(5%, the row's S1 spread over R1–R5); rows whose
  before p50 is under 5 µs also need an absolute increase above 0.5 µs; rows
  whose S1 spread is above 10% also need full separation (every after child
  above every before child). `doc_facade_file_{open,full_text}` on
  `FloatingPictures.doc` are compared mode by mode (the pooled per-child p50s
  are split at their largest gap; each mode needs > 5% and separation).
- Overhead check: S1 "open" rows (`*_open`, `*_open_lifecycle`) use the goal
  budget, unified − leaf ≤ max(5% of leaf, 20 µs); every other row uses the
  S1 rule as a review trigger, with the predecessor's S1 noise.

## Harness fix

S8a (B3) changed `File::text()` for DOCX to write `\n` between paragraphs.
S8b had already merged, and nothing re-ran its DOCX filesystem routes, so at
`b906691f3b`:

- the DOCX leaf full-text legs still timed the leaf's unseparated text
  (`document()?.text()` and `extract_text()`), so the pair compared different
  work (m0a-log S8a, "For S9/S10");
- the leaf signature that `verify_docx` compares with the unified one still
  used the unseparated text, and the fixed full-text oracle concatenated the
  200 paragraphs with nothing.

Every DOCX filesystem route therefore failed its untimed verification. One
sample per route on a binary built at `b906691f3b` (S1 recipe, SHA-256
`2c5c356f…688ce7`, not kept) reproduced it: the `docx_file` child exited 1
with `filesystem child docx_file_unified_eager_open failed ... "DOCX unified
and leaf semantic signatures differ"`, and the other eleven groups exited 0.

`d10dd4c747` fixes the harness, not the library:

- the DOCX leaf full-text legs call `text_with_paragraph_separator("\n")`
  (eager) and `extract_text_with_paragraph_separator("\n")` (source-backed),
  the calls `File::text()` makes; their timed-call descriptions say so;
- `docx_package_signature` uses the separated text;
- `expected_docx_full_text()` joins the pinned payloads with `\n` (10,199
  bytes instead of 10,000).

The leaf eager open keeps its plain `text()` validation pass, because that is
the pass the unified eager owner runs at open. `c475a527df` adds a
`#[cfg(test)]` test that all four DOCX full-text query legs return the
oracle. The measured binary was built at `d10dd4c747`; the test commit does
not change it. Gates on the fix: `cargo fmt --check` and `cargo clippy
--all-targets -- -D warnings` for the harness, and six harness unit tests
(the new one, the two timed-full-text oracle tests, the route round trip, the
source replay and the XLSX leaf text test) pass.

## Results

All times are µs, as the median of per-child p50s. "n (A/B)" counts the
children per leg. "noise (S1)" is S1's spread over R1–R5, and "S10 spread"
is (max − min) / median of this capture's per-child p50s. "explained" rows
are over their trigger and have an entry in `explanations.json` and the
record.

### Facade comparison

| Scenario | Before selector → after selector | Corpus | n (A/B) | before p50 µs | after p50 µs | ratio | noise (S1) | trigger | S10 spread A / B | verdict |
|---|---|---|---|---|---|---|---|---|---|---|
| open | `docx_file_eager_open` → `docx_file_unified_eager_open` | docx-source-backed-media | 2/2 | 10,383 | 10,512 | 1.012 | 0.1% | 5.0% | 3.0% / 2.4% | pass |
| open | `docx_file_source_open` → `docx_file_unified_source_open` | docx-source-backed-media | 2/2 | 344 | 358 | 1.041 | 0.6% | 5.0% | 1.3% / 2.6% | pass |
| list | `docx_file_eager_paragraph_count` → `docx_file_unified_eager_paragraph_count` | docx-source-backed-media | 2/2 | 99.3 | 107 | 1.081 | 1.0% | 5.0% | 0.6% / 0.7% | explained (see record) |
| list | `docx_file_source_paragraph_count` → `docx_file_unified_source_paragraph_count` | docx-source-backed-media | 2/2 | 178 | 180 | 1.007 | 1.4% | 5.0% | 1.6% / 0.9% | pass |
| list | `docx_file_eager_list_paragraphs` → `docx_file_unified_eager_list_paragraphs` | docx-source-backed-media | 2/2 | 143 | 143 | 0.997 | 2.9% | 5.0% | 0.0% / 1.4% | pass |
| list | `docx_file_source_list_paragraphs` → `docx_file_unified_source_list_paragraphs` | docx-source-backed-media | 2/2 | 240 | 241 | 1.001 | 3.4% | 5.0% | 0.2% / 0.1% | pass |
| full-text | `docx_file_eager_full_text` → `docx_file_unified_eager_full_text` | docx-source-backed-media | 2/2 | 58.7 | 68.9 | 1.173 | 1.3% | 5.0% | 0.3% / 0.9% | explained (see record) |
| full-text | `docx_file_source_full_text` → `docx_file_unified_source_full_text` | docx-source-backed-media | 2/2 | 152 | 149 | 0.981 | 0.7% | 5.0% | 6.3% / 0.7% | pass |
| list | `docx_file_eager_open_paragraph_count_lifecycle` → `docx_file_unified_eager_open_paragraph_count_lifecycle` | docx-source-backed-media | 2/2 | 10,456 | 10,389 | 0.994 | 1.8% | 5.0% | 2.5% / 1.6% | pass |
| list | `docx_file_source_open_paragraph_count_lifecycle` → `docx_file_unified_source_open_paragraph_count_lifecycle` | docx-source-backed-media | 2/2 | 540 | 530 | 0.981 | 1.5% | 5.0% | 3.5% / 0.7% | pass |
| full-text | `docx_file_eager_open_full_text_lifecycle` → `docx_file_unified_eager_open_full_text_lifecycle` | docx-source-backed-media | 2/2 | 10,301 | 10,414 | 1.011 | 3.0% | 5.0% | 1.2% / 1.3% | pass |
| full-text | `docx_file_source_open_full_text_lifecycle` → `docx_file_unified_source_open_full_text_lifecycle` | docx-source-backed-media | 2/2 | 505 | 508 | 1.005 | 1.6% | 5.0% | 2.0% / 3.2% | pass |
| open | `pptx_file_eager_open` → `pptx_file_unified_eager_open` | pptx-source-backed-media | 4/4 | 6,995 | 7,054 | 1.009 | 7.3% | 7.3% | 3.4% / 3.9% | pass |
| open | `pptx_file_source_open` → `pptx_file_unified_source_open` | pptx-source-backed-media | 4/4 | 2,781 | 2,870 | 1.032 | 1.5% | 5.0% | 5.8% / 1.6% | pass |
| list | `pptx_file_eager_list_slides` → `pptx_file_unified_eager_list_slides` | pptx-source-backed-media | 4/4 | 5,870 | 5,993 | 1.021 | 1.3% | 5.0% | 0.4% / 0.6% | pass |
| list | `pptx_file_source_list_slides` → `pptx_file_unified_source_list_slides` | pptx-source-backed-media | 4/4 | 6,099 | 6,210 | 1.018 | 0.7% | 5.0% | 0.4% / 0.6% | pass |
| list | `pptx_file_eager_slide_count` → `pptx_file_unified_eager_slide_count` | pptx-source-backed-media | 4/4 | 1.71 | 4.39 | 2.564 | 22.0% | 22.0% | 31.5% / 7.5% | explained (see record) |
| list | `pptx_file_source_slide_count` → `pptx_file_unified_source_slide_count` | pptx-source-backed-media | 4/4 | 1.61 | 4.45 | 2.765 | 55.4% | 55.4% | 12.4% / 2.7% | explained (see record) |
| read-one | `pptx_file_eager_selected_slide` → `pptx_file_unified_eager_selected_slide` | pptx-source-backed-media | 4/4 | 77.0 | 76.4 | 0.992 | 7.0% | 7.0% | 7.3% / 4.2% | pass |
| read-one | `pptx_file_source_selected_slide` → `pptx_file_unified_source_selected_slide` | pptx-source-backed-media | 4/4 | 71.1 | 77.6 | 1.091 | 5.9% | 5.9% | 8.6% / 5.2% | explained (see record) |
| list | `pptx_file_eager_open_slide_count_lifecycle` → `pptx_file_unified_eager_open_slide_count_lifecycle` | pptx-source-backed-media | 4/4 | 7,129 | 7,087 | 0.994 | 5.4% | 5.4% | 1.5% / 4.5% | pass |
| list | `pptx_file_source_open_slide_count_lifecycle` → `pptx_file_unified_source_open_slide_count_lifecycle` | pptx-source-backed-media | 4/4 | 2,819 | 2,896 | 1.027 | 3.5% | 5.0% | 3.7% / 3.2% | pass |
| read-one | `pptx_file_eager_open_selected_slide_lifecycle` → `pptx_file_unified_eager_open_selected_slide_lifecycle` | pptx-source-backed-media | 4/4 | 7,113 | 7,100 | 0.998 | 6.2% | 6.2% | 0.8% / 4.2% | pass |
| read-one | `pptx_file_source_open_selected_slide_lifecycle` → `pptx_file_unified_source_open_selected_slide_lifecycle` | pptx-source-backed-media | 4/4 | 2,862 | 2,949 | 1.030 | 8.5% | 8.5% | 1.3% / 1.5% | pass |
| open | `xlsx_file_open` → `xlsx_file_unified_open` | xlsx-cell-values-medium | 6/6 | 253 | 377 | 1.488 | 19.7% | 19.7% | 17.4% / 15.4% | explained (see record) |
| open | `xlsx_file_open_lifecycle` → `xlsx_file_unified_open_lifecycle` | xlsx-cell-values-medium | 6/6 | 2,648 | 2,801 | 1.058 | 4.1% | 5.0% | 2.5% / 3.7% | explained (see record) |
| read-one | `xlsx_file_selected_cell` → `xlsx_file_unified_selected_cell` | xlsx-cell-values-medium | 6/6 | 2,981 | 3,050 | 1.023 | 0.9% | 5.0% | 0.5% / 0.7% | pass |
| open | `xlsx_bytes_open` → `xlsx_bytes_unified_open` | xlsx-cell-values-medium | 6/6 | 60.4 | 62.3 | 1.032 | 34.1% | 34.1% | 49.5% / 2.0% | pass |
| open | `xlsx_bytes_open_lifecycle` → `xlsx_bytes_unified_open_lifecycle` | xlsx-cell-values-medium | 6/6 | 1,790 | 1,798 | 1.004 | 1.0% | 5.0% | 4.2% / 0.6% | pass |
| read-one | `docx_semantic_one_paragraph_text` → `docx_semantic_unified_one_paragraph_text` | docx-semantic-large | 16/16 | 0.96 | 0.80 | 0.836 | 65.2% | 65.2% | 7.3% / 6.2% | pass |
| read-one | `docx_semantic_one_paragraph_text` → `docx_semantic_unified_one_paragraph_text` | docx-semantic-medium | 16/16 | 0.61 | 0.53 | 0.869 | 5.8% | 5.8% | 7.4% / 11.3% | pass |
| read-one | `docx_semantic_one_paragraph_text` → `docx_semantic_unified_one_paragraph_text` | docx-semantic-tiny | 16/16 | 0.58 | 0.52 | 0.897 | 3.6% | 5.0% | 8.6% / 7.7% | pass |
| open | `doc_facade_file_open` → `doc_file_unified_open` | doc-facade-real-file:FloatingPictures.doc | 16/16 | 180 | 178 | 0.991 | 37.6% | 37.6% | 41.4% / 4.7% | pass |
| open | `doc_facade_file_open` → `doc_file_unified_open` | doc-facade-real-file:documentProperties.doc | 16/16 | 15.7 | 14.4 | 0.922 | 1.8% | 5.0% | 2.8% / 2.3% | pass |
| open | `doc_facade_file_open` → `doc_file_unified_open` | doc-facade-real-file:duplicate-style-names.doc | 16/16 | 11.5 | 10.8 | 0.944 | 1.0% | 5.0% | 1.8% / 1.7% | pass |
| full-text | `doc_facade_file_full_text` → `doc_file_unified_full_text` | doc-facade-real-file:FloatingPictures.doc | 16/16 | 271 | 190 | 0.702 | 31.1% | 31.1% | 33.8% / 4.0% | pass |
| full-text | `doc_facade_file_full_text` → `doc_file_unified_full_text` | doc-facade-real-file:documentProperties.doc | 16/16 | 16.4 | 15.3 | 0.937 | 2.6% | 5.0% | 2.4% / 3.6% | pass |
| full-text | `doc_facade_file_full_text` → `doc_file_unified_full_text` | doc-facade-real-file:duplicate-style-names.doc | 16/16 | 11.3 | 10.9 | 0.971 | 2.0% | 5.0% | 2.9% / 2.0% | pass |
| read-one | `doc_facade_file_one_paragraph` → `doc_file_unified_one_paragraph` | doc-facade-real-file:FloatingPictures.doc | 16/16 | 499 | 422 | 0.847 | 7.4% | 7.4% | 13.7% / 13.3% | pass |
| read-one | `doc_facade_file_one_paragraph` → `doc_file_unified_one_paragraph` | doc-facade-real-file:documentProperties.doc | 16/16 | 17.2 | 16.4 | 0.950 | 2.1% | 5.0% | 2.1% / 3.1% | pass |
| read-one | `doc_facade_file_one_paragraph` → `doc_file_unified_one_paragraph` | doc-facade-real-file:duplicate-style-names.doc | 16/16 | 11.4 | 10.9 | 0.958 | 3.0% | 5.0% | 2.4% / 2.0% | pass |
| open | `ppt_facade_file_open` → `ppt_file_unified_open` | ppt-facade-real-file:SampleShow.ppt | 16/16 | 30.0 | 29.0 | 0.967 | 2.4% | 5.0% | 4.8% / 3.2% | pass |
| open | `ppt_facade_file_open` → `ppt_file_unified_open` | ppt-facade-real-file:ppt_with_png.ppt | 16/16 | 28.1 | 28.1 | 1.000 | 0.4% | 5.0% | 2.7% / 1.8% | pass |
| full-text | `ppt_facade_file_full_text` → `ppt_file_unified_full_text` | ppt-facade-real-file:SampleShow.ppt | 16/16 | 35.2 | 34.9 | 0.991 | 1.6% | 5.0% | 4.5% / 2.9% | pass |
| full-text | `ppt_facade_file_full_text` → `ppt_file_unified_full_text` | ppt-facade-real-file:ppt_with_png.ppt | 16/16 | 32.0 | 32.5 | 1.016 | 1.9% | 5.0% | 2.0% / 1.8% | pass |
| read-one | `ppt_facade_file_one_slide_text` → `ppt_file_unified_one_slide_text` | ppt-facade-real-file:SampleShow.ppt | 16/16 | 35.5 | 35.2 | 0.990 | 3.0% | 5.0% | 3.6% / 2.3% | pass |
| read-one | `ppt_facade_file_one_slide_text` → `ppt_file_unified_one_slide_text` | ppt-facade-real-file:ppt_with_png.ppt | 16/16 | 34.5 | 34.9 | 1.010 | 1.2% | 5.0% | 2.2% / 2.2% | pass |
| open | `rtf_file_open` → `rtf_file_unified_open` | rtf-semantic-plain-large | 16/16 | 1,892 | 3,179 | 1.680 | 8.0% | 8.0% | 7.1% / 5.3% | explained (see record) |
| open | `rtf_file_open` → `rtf_file_unified_open` | rtf-semantic-plain-medium | 16/16 | 51.4 | 52.1 | 1.013 | 3.2% | 5.0% | 2.4% / 2.7% | pass |
| open | `rtf_file_open` → `rtf_file_unified_open` | rtf-semantic-plain-tiny | 16/16 | 7.33 | 7.43 | 1.015 | 3.4% | 5.0% | 2.3% / 3.1% | pass |
| open | `rtf_file_open_lifecycle` → `rtf_file_unified_open_lifecycle` | rtf-semantic-plain-large | 16/16 | 1,937 | 1,964 | 1.014 | 5.2% | 5.2% | 5.4% / 32.2% | pass |
| open | `rtf_file_open_lifecycle` → `rtf_file_unified_open_lifecycle` | rtf-semantic-plain-medium | 16/16 | 52.3 | 42.1 | 0.806 | 3.4% | 5.0% | 2.9% / 2.7% | pass |
| open | `rtf_file_open_lifecycle` → `rtf_file_unified_open_lifecycle` | rtf-semantic-plain-tiny | 16/16 | 7.58 | 7.71 | 1.018 | 2.0% | 5.0% | 2.9% / 2.2% | pass |
| open | `odt_file_eager_open` → `odt_file_unified_eager_open` | odt-source-backed-repeated-text-media | 16/16 | 1,906 | 2,032 | 1.066 | 7.0% | 7.0% | 10.2% / 17.9% | pass |
| open | `odt_file_source_open` → `odt_file_unified_source_open` | odt-source-backed-repeated-text-media | 16/16 | 651 | 645 | 0.990 | 3.4% | 5.0% | 4.2% / 6.9% | pass |
| full-text | `odt_file_eager_open_full_text_lifecycle` → `odt_file_unified_eager_open_full_text_lifecycle` | odt-source-backed-repeated-text-media | 16/16 | 3,249 | 3,156 | 0.971 | 7.2% | 7.2% | 26.6% / 15.2% | pass |
| full-text | `odt_file_source_open_full_text_lifecycle` → `odt_file_unified_source_open_full_text_lifecycle` | odt-source-backed-repeated-text-media | 16/16 | 1,784 | 1,717 | 0.962 | 2.0% | 5.0% | 3.9% / 4.5% | pass |
| open | `odp_file_eager_open` → `odp_file_unified_eager_open` | odp-media-textbox-publication | 8/8 | 163 | 133 | 0.815 | 3.6% | 5.0% | 25.5% / 14.2% | pass |
| open | `odp_file_source_open` → `odp_file_unified_source_open` | odp-media-textbox-publication | 8/8 | 102 | 43.3 | 0.424 | 12.1% | 12.1% | 35.8% / 1.5% | pass |
| read-one | `odp_file_eager_selected_slide` → `odp_file_unified_eager_selected_slide` | odp-media-textbox-publication | 8/8 | 79.2 | 56.5 | 0.713 | 5.9% | 5.9% | 18.6% / 10.4% | pass |
| read-one | `odp_file_source_selected_slide` → `odp_file_unified_source_selected_slide` | odp-media-textbox-publication | 8/8 | 73.7 | 49.2 | 0.668 | 9.0% | 9.0% | 16.1% / 2.1% | pass |
| open | `ods_file_eager_open` → `ods_file_unified_eager_open` | ods-media-publication | 8/8 | 2,578 | 2,442 | 0.948 | 5.2% | 5.2% | 118.5% / 126.5% | pass |
| open | `ods_file_source_open` → `ods_file_unified_source_open` | ods-media-publication | 8/8 | 2,012 | 1,871 | 0.930 | 7.9% | 7.9% | 152.7% / 163.1% | pass |
| detection | `detect_odt_bytes` → `detect_odt_unified_bytes` | odt-bytes-large | 16/16 | 1.48 | 1.47 | 0.995 | 2.7% | 5.0% | 5.4% / 6.8% | pass |
| detection | `detect_odt_bytes` → `detect_odt_unified_bytes` | odt-bytes-medium | 16/16 | 1.46 | 1.45 | 0.997 | 2.7% | 5.0% | 5.5% / 6.2% | pass |
| detection | `detect_odt_bytes` → `detect_odt_unified_bytes` | odt-bytes-tiny | 16/16 | 1.45 | 1.48 | 1.015 | 0.3% | 5.0% | 4.1% / 4.7% | pass |
| detection | `detect_odt_reader` → `detect_odt_unified_reader` | odt-reader-large | 16/16 | 7.07 | 7.19 | 1.017 | 0.4% | 5.0% | 2.5% / 2.4% | pass |
| detection | `detect_odt_reader` → `detect_odt_unified_reader` | odt-reader-medium | 16/16 | 6.41 | 6.57 | 1.025 | 1.9% | 5.0% | 3.7% / 5.2% | pass |
| detection | `detect_odt_reader` → `detect_odt_unified_reader` | odt-reader-tiny | 16/16 | 6.38 | 6.51 | 1.020 | 2.2% | 5.0% | 3.4% / 5.5% | pass |
| detection | `detect_odt_polyglot` → `detect_odt_unified_polyglot` | odt-polyglot-large | 16/16 | 14.2 | 14.5 | 1.027 | 2.3% | 5.0% | 5.0% / 1.5% | pass |
| detection | `detect_odt_polyglot` → `detect_odt_unified_polyglot` | odt-polyglot-medium | 16/16 | 14.1 | 14.5 | 1.026 | 1.4% | 5.0% | 3.8% / 2.0% | pass |
| detection | `detect_odt_polyglot` → `detect_odt_unified_polyglot` | odt-polyglot-tiny | 16/16 | 14.3 | 14.6 | 1.017 | 2.1% | 5.0% | 2.2% / 3.6% | pass |
| detection | `detect_odt_catalog_alias` → `detect_odt_unified_catalog_alias` | odt-catalog-alias-large | 16/16 | 12.0 | 12.4 | 1.029 | 3.0% | 5.0% | 5.3% / 2.0% | pass |
| detection | `detect_odt_catalog_alias` → `detect_odt_unified_catalog_alias` | odt-catalog-alias-medium | 16/16 | 12.0 | 12.3 | 1.026 | 2.7% | 5.0% | 4.4% / 3.3% | pass |
| detection | `detect_odt_catalog_alias` → `detect_odt_unified_catalog_alias` | odt-catalog-alias-tiny | 16/16 | 12.0 | 12.4 | 1.027 | 1.7% | 5.0% | 5.5% / 3.0% | pass |
| detection | `detect_ods_bytes` → `detect_ods_unified_bytes` | ods-bytes-large | 16/16 | 1.11 | 1.12 | 1.009 | 1.8% | 5.0% | 4.5% / 5.4% | pass |
| detection | `detect_ods_bytes` → `detect_ods_unified_bytes` | ods-bytes-medium | 16/16 | 1.08 | 1.08 | 1.000 | 1.9% | 5.0% | 2.8% / 4.6% | pass |
| detection | `detect_ods_bytes` → `detect_ods_unified_bytes` | ods-bytes-tiny | 16/16 | 1.07 | 1.07 | 0.995 | 0.9% | 5.0% | 2.8% / 3.7% | pass |
| detection | `detect_ods_reader` → `detect_ods_unified_reader` | ods-reader-large | 16/16 | 7.77 | 7.71 | 0.993 | 0.5% | 5.0% | 10.4% / 15.8% | pass |
| detection | `detect_ods_reader` → `detect_ods_unified_reader` | ods-reader-medium | 16/16 | 5.19 | 5.20 | 1.001 | 1.1% | 5.0% | 3.5% / 4.8% | pass |
| detection | `detect_ods_reader` → `detect_ods_unified_reader` | ods-reader-tiny | 16/16 | 4.97 | 5.07 | 1.019 | 2.2% | 5.0% | 4.5% / 5.0% | pass |
| detection | `detect_ods_polyglot` → `detect_ods_unified_polyglot` | ods-polyglot-large | 16/16 | 12.6 | 12.9 | 1.020 | 1.0% | 5.0% | 3.1% / 2.1% | pass |
| detection | `detect_ods_polyglot` → `detect_ods_unified_polyglot` | ods-polyglot-medium | 16/16 | 12.4 | 12.8 | 1.031 | 1.9% | 5.0% | 2.1% / 3.3% | pass |
| detection | `detect_ods_polyglot` → `detect_ods_unified_polyglot` | ods-polyglot-tiny | 16/16 | 12.3 | 12.8 | 1.041 | 2.3% | 5.0% | 2.2% / 1.8% | pass |
| detection | `detect_ods_catalog_alias` → `detect_ods_unified_catalog_alias` | ods-catalog-alias-large | 16/16 | 10.6 | 10.9 | 1.025 | 2.6% | 5.0% | 3.3% / 3.2% | pass |
| detection | `detect_ods_catalog_alias` → `detect_ods_unified_catalog_alias` | ods-catalog-alias-medium | 16/16 | 10.5 | 10.8 | 1.032 | 3.4% | 5.0% | 2.5% / 3.3% | pass |
| detection | `detect_ods_catalog_alias` → `detect_ods_unified_catalog_alias` | ods-catalog-alias-tiny | 16/16 | 10.4 | 10.8 | 1.040 | 2.4% | 5.0% | 2.7% / 3.1% | pass |
| detection | `detect_odp_bytes` → `detect_odp_unified_bytes` | odp-bytes-large | 16/16 | 1.52 | 1.52 | 1.002 | 1.3% | 5.0% | 6.6% / 2.6% | pass |
| detection | `detect_odp_bytes` → `detect_odp_unified_bytes` | odp-bytes-medium | 16/16 | 1.52 | 1.53 | 1.005 | 3.3% | 5.0% | 5.6% / 5.9% | pass |
| detection | `detect_odp_bytes` → `detect_odp_unified_bytes` | odp-bytes-tiny | 16/16 | 1.50 | 1.49 | 0.993 | 2.7% | 5.0% | 3.3% / 4.7% | pass |
| detection | `detect_odp_reader` → `detect_odp_unified_reader` | odp-reader-large | 16/16 | 6.63 | 6.66 | 1.004 | 1.4% | 5.0% | 4.1% / 3.2% | pass |
| detection | `detect_odp_reader` → `detect_odp_unified_reader` | odp-reader-medium | 16/16 | 6.57 | 6.68 | 1.018 | 0.6% | 5.0% | 5.3% / 4.2% | pass |
| detection | `detect_odp_reader` → `detect_odp_unified_reader` | odp-reader-tiny | 16/16 | 6.46 | 6.64 | 1.027 | 2.6% | 5.0% | 4.2% / 4.5% | pass |
| detection | `detect_odp_polyglot` → `detect_odp_unified_polyglot` | odp-polyglot-large | 16/16 | 14.5 | 14.6 | 1.009 | 1.5% | 5.0% | 3.6% / 2.4% | pass |
| detection | `detect_odp_polyglot` → `detect_odp_unified_polyglot` | odp-polyglot-medium | 16/16 | 14.2 | 14.6 | 1.026 | 2.9% | 5.0% | 3.0% / 2.5% | pass |
| detection | `detect_odp_polyglot` → `detect_odp_unified_polyglot` | odp-polyglot-tiny | 16/16 | 14.4 | 14.7 | 1.018 | 3.4% | 5.0% | 3.1% / 3.0% | pass |
| detection | `detect_odp_catalog_alias` → `detect_odp_unified_catalog_alias` | odp-catalog-alias-large | 16/16 | 12.3 | 12.5 | 1.014 | 3.3% | 5.0% | 4.2% / 2.5% | pass |
| detection | `detect_odp_catalog_alias` → `detect_odp_unified_catalog_alias` | odp-catalog-alias-medium | 16/16 | 12.2 | 12.5 | 1.021 | 2.0% | 5.0% | 4.6% / 1.6% | pass |
| detection | `detect_odp_catalog_alias` → `detect_odp_unified_catalog_alias` | odp-catalog-alias-tiny | 16/16 | 12.3 | 12.5 | 1.017 | 4.5% | 5.0% | 3.9% / 2.8% | pass |

The bimodal rows (`doc_facade_file_{open,full_text}` on `FloatingPictures.doc`) were compared mode by mode, as S1 required. The pre-M0a binary was in the high mode in 5 of 16 children (open) and 9 of 16 (full text); the post-S9 binary was in the low mode in all 16. The low-mode ratios are 1.001 and 0.999 (`modes` in `analysis.json`). That is why the pooled full-text ratio is 0.70 and is not a claim.

### Kept predecessors (same selector, both binaries)

| Scenario | Before selector → after selector | Corpus | n (A/B) | before p50 µs | after p50 µs | ratio | noise (S1) | trigger | S10 spread A / B | verdict |
|---|---|---|---|---|---|---|---|---|---|---|
| detection | `detect_odp_bytes` | odp-bytes-large | 16/16 | 1.52 | 1.52 | 1.002 | 1.3% | 5.0% | 6.6% / 3.9% | pass |
| detection | `detect_odp_bytes` | odp-bytes-medium | 16/16 | 1.52 | 1.50 | 0.990 | 3.3% | 5.0% | 5.6% / 6.0% | pass |
| detection | `detect_odp_bytes` | odp-bytes-tiny | 16/16 | 1.50 | 1.49 | 0.992 | 2.7% | 5.0% | 3.3% / 3.7% | pass |
| detection | `detect_odp_catalog_alias` | odp-catalog-alias-large | 16/16 | 12.3 | 12.5 | 1.018 | 3.3% | 5.0% | 4.2% / 4.2% | pass |
| detection | `detect_odp_catalog_alias` | odp-catalog-alias-medium | 16/16 | 12.2 | 12.4 | 1.018 | 2.0% | 5.0% | 4.6% / 3.6% | pass |
| detection | `detect_odp_catalog_alias` | odp-catalog-alias-tiny | 16/16 | 12.3 | 12.5 | 1.015 | 4.5% | 5.0% | 3.9% / 2.6% | pass |
| detection | `detect_odp_polyglot` | odp-polyglot-large | 16/16 | 14.5 | 14.6 | 1.009 | 1.5% | 5.0% | 3.6% / 2.7% | pass |
| detection | `detect_odp_polyglot` | odp-polyglot-medium | 16/16 | 14.2 | 14.5 | 1.022 | 2.9% | 5.0% | 3.0% / 2.3% | pass |
| detection | `detect_odp_polyglot` | odp-polyglot-tiny | 16/16 | 14.4 | 14.6 | 1.014 | 3.4% | 5.0% | 3.1% / 2.3% | pass |
| detection | `detect_odp_reader` | odp-reader-large | 16/16 | 6.63 | 6.62 | 0.998 | 1.4% | 5.0% | 4.1% / 5.1% | pass |
| detection | `detect_odp_reader` | odp-reader-medium | 16/16 | 6.57 | 6.59 | 1.003 | 0.6% | 5.0% | 5.3% / 3.6% | pass |
| detection | `detect_odp_reader` | odp-reader-tiny | 16/16 | 6.46 | 6.50 | 1.005 | 2.6% | 5.0% | 4.2% / 4.2% | pass |
| detection | `detect_ods_bytes` | ods-bytes-large | 16/16 | 1.11 | 1.12 | 1.009 | 1.8% | 5.0% | 4.5% / 4.9% | pass |
| detection | `detect_ods_bytes` | ods-bytes-medium | 16/16 | 1.08 | 1.07 | 0.991 | 1.9% | 5.0% | 2.8% / 3.3% | pass |
| detection | `detect_ods_bytes` | ods-bytes-tiny | 16/16 | 1.07 | 1.08 | 1.002 | 0.9% | 5.0% | 2.8% / 4.2% | pass |
| detection | `detect_ods_catalog_alias` | ods-catalog-alias-large | 16/16 | 10.6 | 10.8 | 1.017 | 2.6% | 5.0% | 3.3% / 3.0% | pass |
| detection | `detect_ods_catalog_alias` | ods-catalog-alias-medium | 16/16 | 10.5 | 10.7 | 1.022 | 3.4% | 5.0% | 2.5% / 2.8% | pass |
| detection | `detect_ods_catalog_alias` | ods-catalog-alias-tiny | 16/16 | 10.4 | 10.6 | 1.022 | 2.4% | 5.0% | 2.7% / 2.4% | pass |
| detection | `detect_ods_polyglot` | ods-polyglot-large | 16/16 | 12.6 | 12.8 | 1.016 | 1.0% | 5.0% | 3.1% / 2.7% | pass |
| detection | `detect_ods_polyglot` | ods-polyglot-medium | 16/16 | 12.4 | 12.7 | 1.027 | 1.9% | 5.0% | 2.1% / 2.7% | pass |
| detection | `detect_ods_polyglot` | ods-polyglot-tiny | 16/16 | 12.3 | 12.5 | 1.018 | 2.3% | 5.0% | 2.2% / 3.2% | pass |
| detection | `detect_ods_reader` | ods-reader-large | 16/16 | 7.77 | 7.75 | 0.996 | 0.5% | 5.0% | 10.4% / 9.6% | pass |
| detection | `detect_ods_reader` | ods-reader-medium | 16/16 | 5.19 | 5.15 | 0.992 | 1.1% | 5.0% | 3.5% / 4.4% | pass |
| detection | `detect_ods_reader` | ods-reader-tiny | 16/16 | 4.97 | 5.01 | 1.008 | 2.2% | 5.0% | 4.5% / 3.0% | pass |
| detection | `detect_odt_bytes` | odt-bytes-large | 16/16 | 1.48 | 1.47 | 0.997 | 2.7% | 5.0% | 5.4% / 8.2% | pass |
| detection | `detect_odt_bytes` | odt-bytes-medium | 16/16 | 1.46 | 1.45 | 0.997 | 2.7% | 5.0% | 5.5% / 3.4% | pass |
| detection | `detect_odt_bytes` | odt-bytes-tiny | 16/16 | 1.45 | 1.45 | 0.998 | 0.3% | 5.0% | 4.1% / 3.4% | pass |
| detection | `detect_odt_catalog_alias` | odt-catalog-alias-large | 16/16 | 12.0 | 12.4 | 1.025 | 3.0% | 5.0% | 5.3% / 3.4% | pass |
| detection | `detect_odt_catalog_alias` | odt-catalog-alias-medium | 16/16 | 12.0 | 12.3 | 1.022 | 2.7% | 5.0% | 4.4% / 3.4% | pass |
| detection | `detect_odt_catalog_alias` | odt-catalog-alias-tiny | 16/16 | 12.0 | 12.2 | 1.016 | 1.7% | 5.0% | 5.5% / 3.5% | pass |
| detection | `detect_odt_polyglot` | odt-polyglot-large | 16/16 | 14.2 | 14.5 | 1.024 | 2.3% | 5.0% | 5.0% / 3.1% | pass |
| detection | `detect_odt_polyglot` | odt-polyglot-medium | 16/16 | 14.1 | 14.4 | 1.022 | 1.4% | 5.0% | 3.8% / 2.0% | pass |
| detection | `detect_odt_polyglot` | odt-polyglot-tiny | 16/16 | 14.3 | 14.4 | 1.009 | 2.1% | 5.0% | 2.2% / 2.6% | pass |
| detection | `detect_odt_reader` | odt-reader-large | 16/16 | 7.07 | 7.07 | 1.000 | 0.4% | 5.0% | 2.5% / 2.8% | pass |
| detection | `detect_odt_reader` | odt-reader-medium | 16/16 | 6.41 | 6.43 | 1.005 | 1.9% | 5.0% | 3.7% / 3.7% | pass |
| detection | `detect_odt_reader` | odt-reader-tiny | 16/16 | 6.38 | 6.35 | 0.995 | 2.2% | 5.0% | 3.4% / 4.4% | pass |
| read-one | `docx_producer_source_selected_paragraph` | docx-producer-medium | 16/16 | 106 | 106 | 0.997 | 4.6% | 5.0% | 6.5% / 4.9% | pass |
| read-one | `docx_semantic_one_paragraph` | docx-semantic-large | 16/16 | 3,186 | 3,174 | 0.996 | 1.1% | 5.0% | 7.2% / 3.0% | pass |
| read-one | `docx_semantic_one_paragraph` | docx-semantic-medium | 16/16 | 65.3 | 64.9 | 0.995 | 2.0% | 5.0% | 8.4% / 3.1% | pass |
| read-one | `docx_semantic_one_paragraph` | docx-semantic-tiny | 16/16 | 9.58 | 9.52 | 0.993 | 1.7% | 5.0% | 5.6% / 2.6% | pass |
| read-one | `docx_semantic_one_paragraph_text` | docx-semantic-large | 16/16 | 0.96 | 0.92 | 0.958 | 65.2% | 65.2% | 7.3% / 9.8% | pass |
| read-one | `docx_semantic_one_paragraph_text` | docx-semantic-medium | 16/16 | 0.61 | 0.60 | 0.984 | 5.8% | 5.8% | 7.4% / 3.3% | pass |
| read-one | `docx_semantic_one_paragraph_text` | docx-semantic-tiny | 16/16 | 0.58 | 0.57 | 0.983 | 3.6% | 5.0% | 8.6% / 3.5% | pass |
| read-one | `ods_file_eager_selected_cell` | ods-media-publication | 8/8 | 0.60 | 0.52 | 0.867 | 25.5% | 25.5% | 65.0% / 42.3% | pass |
| read-one | `ods_file_eager_selected_media` | ods-media-publication | 8/8 | 267 | 270 | 1.012 | 6.9% | 6.9% | 2.1% / 3.2% | pass |
| read-one | `ods_file_source_selected_cell` | ods-media-publication | 8/8 | 1.36 | 1.38 | 1.009 | 15.8% | 15.8% | 27.5% / 10.2% | pass |
| read-one | `ods_file_source_selected_media` | ods-media-publication | 8/8 | 412 | 420 | 1.020 | 5.5% | 5.5% | 5.1% / 8.7% | pass |

### Budget-only successors

These successors time `Package::document()` as well, because the unified view keeps no document across calls (S8b). They are compared with their leaf twins only.

| Scenario | Before selector → after selector | Corpus | n (A/B) | before p50 µs | after p50 µs | ratio | noise (S1) | trigger | S10 spread A / B | verdict |
|---|---|---|---|---|---|---|---|---|---|---|
| read-one | `docx_semantic_one_paragraph` → `docx_semantic_unified_one_paragraph` | docx-semantic-large | 16/16 | 3,186 | 3,187 | 1.000 | 1.1% | 5.0% | 7.2% / 1.7% | budget-only |
| read-one | `docx_semantic_one_paragraph` → `docx_semantic_unified_one_paragraph` | docx-semantic-medium | 16/16 | 65.3 | 66.3 | 1.015 | 2.0% | 5.0% | 8.4% / 1.6% | budget-only |
| read-one | `docx_semantic_one_paragraph` → `docx_semantic_unified_one_paragraph` | docx-semantic-tiny | 16/16 | 9.58 | 10.6 | 1.104 | 1.7% | 5.0% | 5.6% / 1.4% | budget-only |
| read-one | `docx_producer_source_selected_paragraph` → `docx_producer_unified_source_selected_paragraph` | docx-producer-medium | 16/16 | 106 | 261 | 2.459 | 4.6% | 5.0% | 6.5% / 2.4% | budget-only |

### Unified vs leaf (post-S9 binary)

| Scenario | Unified selector (vs `_leaf_` twin) | Corpus | n (L/U) | leaf p50 µs | unified p50 µs | ratio | Δ µs | rule | verdict |
|---|---|---|---|---|---|---|---|---|---|
| open | `docx_file_unified_eager_open` | docx-source-backed-media | 2/2 | 10,387 | 10,512 | 1.012 | +124.76 | open budget max(5%, 20 µs) | within budget |
| open | `docx_file_unified_source_open` | docx-source-backed-media | 2/2 | 324 | 358 | 1.108 | +34.91 | open budget max(5%, 20 µs) | over budget, explained (see record) |
| list | `docx_file_unified_eager_paragraph_count` | docx-source-backed-media | 2/2 | 104 | 107 | 1.031 | +3.28 | review trigger 5.0% | within trigger |
| list | `docx_file_unified_source_paragraph_count` | docx-source-backed-media | 2/2 | 179 | 180 | 1.001 | +0.16 | review trigger 5.0% | within trigger |
| list | `docx_file_unified_eager_list_paragraphs` | docx-source-backed-media | 2/2 | 109 | 143 | 1.315 | +34.22 | review trigger 5.0% | review, explained (see record) |
| list | `docx_file_unified_source_list_paragraphs` | docx-source-backed-media | 2/2 | 184 | 241 | 1.305 | +56.24 | review trigger 5.0% | review, explained (see record) |
| full-text | `docx_file_unified_eager_full_text` | docx-source-backed-media | 2/2 | 67.6 | 68.9 | 1.020 | +1.32 | review trigger 5.0% | within trigger |
| full-text | `docx_file_unified_source_full_text` | docx-source-backed-media | 2/2 | 150 | 149 | 0.998 | -0.35 | review trigger 5.0% | within trigger |
| list | `docx_file_unified_eager_open_paragraph_count_lifecycle` | docx-source-backed-media | 2/2 | 10,461 | 10,389 | 0.993 | -72.12 | review trigger 5.0% | within trigger |
| list | `docx_file_unified_source_open_paragraph_count_lifecycle` | docx-source-backed-media | 2/2 | 505 | 530 | 1.049 | +24.85 | review trigger 5.0% | within trigger |
| full-text | `docx_file_unified_eager_open_full_text_lifecycle` | docx-source-backed-media | 2/2 | 10,425 | 10,414 | 0.999 | -10.66 | review trigger 5.0% | within trigger |
| full-text | `docx_file_unified_source_open_full_text_lifecycle` | docx-source-backed-media | 2/2 | 478 | 508 | 1.063 | +30.19 | review trigger 5.0% | review, explained (see record) |
| open | `pptx_file_unified_eager_open` | pptx-source-backed-media | 4/4 | 6,928 | 7,054 | 1.018 | +126.08 | open budget max(5%, 20 µs) | within budget |
| open | `pptx_file_unified_source_open` | pptx-source-backed-media | 4/4 | 2,825 | 2,870 | 1.016 | +45.02 | open budget max(5%, 20 µs) | within budget |
| list | `pptx_file_unified_eager_list_slides` | pptx-source-backed-media | 4/4 | 6,021 | 5,993 | 0.995 | -27.72 | review trigger 5.0% | within trigger |
| list | `pptx_file_unified_source_list_slides` | pptx-source-backed-media | 4/4 | 6,261 | 6,210 | 0.992 | -50.91 | review trigger 5.0% | within trigger |
| list | `pptx_file_unified_eager_slide_count` | pptx-source-backed-media | 4/4 | 3.86 | 4.39 | 1.137 | +0.53 | review trigger 22.0% | within trigger |
| list | `pptx_file_unified_source_slide_count` | pptx-source-backed-media | 4/4 | 4.01 | 4.45 | 1.108 | +0.43 | review trigger 55.4% | within trigger |
| read-one | `pptx_file_unified_eager_selected_slide` | pptx-source-backed-media | 4/4 | 73.4 | 76.4 | 1.041 | +2.98 | review trigger 7.0% | within trigger |
| read-one | `pptx_file_unified_source_selected_slide` | pptx-source-backed-media | 4/4 | 75.9 | 77.6 | 1.022 | +1.68 | review trigger 5.9% | within trigger |
| list | `pptx_file_unified_eager_open_slide_count_lifecycle` | pptx-source-backed-media | 4/4 | 6,965 | 7,087 | 1.018 | +122.06 | review trigger 5.4% | within trigger |
| list | `pptx_file_unified_source_open_slide_count_lifecycle` | pptx-source-backed-media | 4/4 | 2,846 | 2,896 | 1.018 | +50.51 | review trigger 5.0% | within trigger |
| read-one | `pptx_file_unified_eager_open_selected_slide_lifecycle` | pptx-source-backed-media | 4/4 | 6,990 | 7,100 | 1.016 | +109.91 | review trigger 6.2% | within trigger |
| read-one | `pptx_file_unified_source_open_selected_slide_lifecycle` | pptx-source-backed-media | 4/4 | 2,895 | 2,949 | 1.019 | +53.92 | review trigger 8.5% | within trigger |
| open | `xlsx_file_unified_open` | xlsx-cell-values-medium | 6/6 | 352 | 377 | 1.070 | +24.64 | open budget max(5%, 20 µs) | over budget, explained (see record) |
| open | `xlsx_file_unified_open_lifecycle` | xlsx-cell-values-medium | 6/6 | 2,688 | 2,801 | 1.042 | +113.22 | open budget max(5%, 20 µs) | within budget |
| read-one | `xlsx_file_unified_selected_cell` | xlsx-cell-values-medium | 6/6 | 3,045 | 3,050 | 1.001 | +4.15 | review trigger 5.0% | within trigger |
| open | `xlsx_bytes_unified_open` | xlsx-cell-values-medium | 6/6 | 61.5 | 62.3 | 1.014 | +0.83 | open budget max(5%, 20 µs) | within budget |
| open | `xlsx_bytes_unified_open_lifecycle` | xlsx-cell-values-medium | 6/6 | 1,783 | 1,798 | 1.008 | +14.33 | open budget max(5%, 20 µs) | within budget |
| read-one | `docx_semantic_unified_one_paragraph` | docx-semantic-large | 16/16 | 3,204 | 3,187 | 0.995 | -17.32 | review trigger 5.0% | within trigger |
| read-one | `docx_semantic_unified_one_paragraph` | docx-semantic-medium | 16/16 | 66.3 | 66.3 | 1.000 | -0.02 | review trigger 5.0% | within trigger |
| read-one | `docx_semantic_unified_one_paragraph` | docx-semantic-tiny | 16/16 | 10.7 | 10.6 | 0.991 | -0.09 | review trigger 5.0% | within trigger |
| read-one | `docx_semantic_unified_one_paragraph_text` | docx-semantic-large | 16/16 | 0.77 | 0.80 | 1.039 | +0.03 | review trigger 65.2% | within trigger |
| read-one | `docx_semantic_unified_one_paragraph_text` | docx-semantic-medium | 16/16 | 0.52 | 0.53 | 1.019 | +0.01 | review trigger 5.8% | within trigger |
| read-one | `docx_semantic_unified_one_paragraph_text` | docx-semantic-tiny | 16/16 | 0.51 | 0.52 | 1.020 | +0.01 | review trigger 5.0% | within trigger |
| read-one | `docx_producer_unified_source_selected_paragraph` | docx-producer-medium | 16/16 | 263 | 261 | 0.993 | -1.85 | review trigger 5.0% | within trigger |
| open | `doc_file_unified_open` | doc-facade-real-file:FloatingPictures.doc | 16/16 | 175 | 178 | 1.019 | +3.31 | open budget max(5%, 20 µs) | within budget |
| open | `doc_file_unified_open` | doc-facade-real-file:documentProperties.doc | 16/16 | 13.3 | 14.4 | 1.086 | +1.14 | open budget max(5%, 20 µs) | within budget |
| open | `doc_file_unified_open` | doc-facade-real-file:duplicate-style-names.doc | 16/16 | 9.26 | 10.8 | 1.170 | +1.57 | open budget max(5%, 20 µs) | within budget |
| full-text | `doc_file_unified_full_text` | doc-facade-real-file:FloatingPictures.doc | 16/16 | 187 | 190 | 1.018 | +3.29 | review trigger 31.1% | within trigger |
| full-text | `doc_file_unified_full_text` | doc-facade-real-file:documentProperties.doc | 16/16 | 14.2 | 15.3 | 1.084 | +1.19 | review trigger 5.0% | review, explained (see record) |
| full-text | `doc_file_unified_full_text` | doc-facade-real-file:duplicate-style-names.doc | 16/16 | 9.37 | 10.9 | 1.169 | +1.58 | review trigger 5.0% | review, explained (see record) |
| read-one | `doc_file_unified_one_paragraph` | doc-facade-real-file:FloatingPictures.doc | 16/16 | 415 | 422 | 1.016 | +6.73 | review trigger 7.4% | within trigger |
| read-one | `doc_file_unified_one_paragraph` | doc-facade-real-file:documentProperties.doc | 16/16 | 15.1 | 16.4 | 1.084 | +1.26 | review trigger 5.0% | review, explained (see record) |
| read-one | `doc_file_unified_one_paragraph` | doc-facade-real-file:duplicate-style-names.doc | 16/16 | 9.35 | 10.9 | 1.165 | +1.54 | review trigger 5.0% | review, explained (see record) |
| open | `ppt_file_unified_open` | ppt-facade-real-file:SampleShow.ppt | 16/16 | 28.7 | 29.0 | 1.012 | +0.34 | open budget max(5%, 20 µs) | within budget |
| open | `ppt_file_unified_open` | ppt-facade-real-file:ppt_with_png.ppt | 16/16 | 27.4 | 28.1 | 1.025 | +0.69 | open budget max(5%, 20 µs) | within budget |
| full-text | `ppt_file_unified_full_text` | ppt-facade-real-file:SampleShow.ppt | 16/16 | 34.3 | 34.9 | 1.016 | +0.56 | review trigger 5.0% | within trigger |
| full-text | `ppt_file_unified_full_text` | ppt-facade-real-file:ppt_with_png.ppt | 16/16 | 31.7 | 32.5 | 1.025 | +0.78 | review trigger 5.0% | within trigger |
| read-one | `ppt_file_unified_one_slide_text` | ppt-facade-real-file:SampleShow.ppt | 16/16 | 34.5 | 35.2 | 1.019 | +0.65 | review trigger 5.0% | within trigger |
| read-one | `ppt_file_unified_one_slide_text` | ppt-facade-real-file:ppt_with_png.ppt | 16/16 | 33.9 | 34.9 | 1.027 | +0.93 | review trigger 5.0% | within trigger |
| open | `rtf_file_unified_open` | rtf-semantic-plain-large | 16/16 | 1,825 | 3,179 | 1.742 | +1354.67 | open budget max(5%, 20 µs) | over budget, explained (see record) |
| open | `rtf_file_unified_open` | rtf-semantic-plain-medium | 16/16 | 47.8 | 52.1 | 1.091 | +4.33 | open budget max(5%, 20 µs) | within budget |
| open | `rtf_file_unified_open` | rtf-semantic-plain-tiny | 16/16 | 7.26 | 7.43 | 1.024 | +0.18 | open budget max(5%, 20 µs) | within budget |
| open | `rtf_file_unified_open_lifecycle` | rtf-semantic-plain-large | 16/16 | 1,862 | 1,964 | 1.055 | +102.30 | open budget max(5%, 20 µs) | over budget, explained (see record) |
| open | `rtf_file_unified_open_lifecycle` | rtf-semantic-plain-medium | 16/16 | 51.1 | 42.1 | 0.826 | -8.91 | open budget max(5%, 20 µs) | within budget |
| open | `rtf_file_unified_open_lifecycle` | rtf-semantic-plain-tiny | 16/16 | 7.33 | 7.71 | 1.051 | +0.38 | open budget max(5%, 20 µs) | within budget |
| open | `odt_file_unified_eager_open` | odt-source-backed-repeated-text-media | 16/16 | 2,024 | 2,032 | 1.004 | +8.19 | open budget max(5%, 20 µs) | within budget |
| open | `odt_file_unified_source_open` | odt-source-backed-repeated-text-media | 16/16 | 650 | 645 | 0.992 | -4.92 | open budget max(5%, 20 µs) | within budget |
| full-text | `odt_file_unified_eager_open_full_text_lifecycle` | odt-source-backed-repeated-text-media | 16/16 | 3,127 | 3,156 | 1.009 | +29.49 | review trigger 7.2% | within trigger |
| full-text | `odt_file_unified_source_open_full_text_lifecycle` | odt-source-backed-repeated-text-media | 16/16 | 1,719 | 1,717 | 0.999 | -2.01 | review trigger 5.0% | within trigger |
| open | `odp_file_unified_eager_open` | odp-media-textbox-publication | 8/8 | 131 | 133 | 1.016 | +2.06 | open budget max(5%, 20 µs) | within budget |
| open | `odp_file_unified_source_open` | odp-media-textbox-publication | 8/8 | 38.5 | 43.3 | 1.126 | +4.86 | open budget max(5%, 20 µs) | within budget |
| read-one | `odp_file_unified_eager_selected_slide` | odp-media-textbox-publication | 8/8 | 56.8 | 56.5 | 0.995 | -0.31 | review trigger 5.9% | within trigger |
| read-one | `odp_file_unified_source_selected_slide` | odp-media-textbox-publication | 8/8 | 49.3 | 49.2 | 0.998 | -0.10 | review trigger 9.0% | within trigger |
| open | `ods_file_unified_eager_open` | ods-media-publication | 8/8 | 3,937 | 2,442 | 0.620 | -1494.33 | open budget max(5%, 20 µs) | within budget |
| open | `ods_file_unified_source_open` | ods-media-publication | 8/8 | 3,371 | 1,871 | 0.555 | -1500.07 | open budget max(5%, 20 µs) | within budget |
| detection | `detect_odt_unified_bytes` | odt-bytes-large | 16/16 | 0.85 | 1.47 | 1.726 | +0.62 | review trigger 5.0% | review, explained (see record) |
| detection | `detect_odt_unified_bytes` | odt-bytes-medium | 16/16 | 0.85 | 1.45 | 1.706 | +0.60 | review trigger 5.0% | review, explained (see record) |
| detection | `detect_odt_unified_bytes` | odt-bytes-tiny | 16/16 | 0.83 | 1.48 | 1.777 | +0.64 | review trigger 5.0% | review, explained (see record) |
| detection | `detect_odt_unified_reader` | odt-reader-large | 16/16 | 1.82 | 7.19 | 3.951 | +5.37 | review trigger 5.0% | review, explained (see record) |
| detection | `detect_odt_unified_reader` | odt-reader-medium | 16/16 | 1.25 | 6.57 | 5.252 | +5.32 | review trigger 5.0% | review, explained (see record) |
| detection | `detect_odt_unified_reader` | odt-reader-tiny | 16/16 | 1.23 | 6.51 | 5.314 | +5.29 | review trigger 5.0% | review, explained (see record) |
| detection | `detect_ods_unified_bytes` | ods-bytes-large | 16/16 | 0.61 | 1.12 | 1.836 | +0.51 | review trigger 5.0% | review, explained (see record) |
| detection | `detect_ods_unified_bytes` | ods-bytes-medium | 16/16 | 0.60 | 1.08 | 1.800 | +0.48 | review trigger 5.0% | within trigger |
| detection | `detect_ods_unified_bytes` | ods-bytes-tiny | 16/16 | 0.60 | 1.07 | 1.783 | +0.47 | review trigger 5.0% | within trigger |
| detection | `detect_ods_unified_reader` | ods-reader-large | 16/16 | 3.65 | 7.71 | 2.114 | +4.07 | review trigger 5.0% | review, explained (see record) |
| detection | `detect_ods_unified_reader` | ods-reader-medium | 16/16 | 1.08 | 5.20 | 4.812 | +4.12 | review trigger 5.0% | review, explained (see record) |
| detection | `detect_ods_unified_reader` | ods-reader-tiny | 16/16 | 0.95 | 5.07 | 5.337 | +4.12 | review trigger 5.0% | review, explained (see record) |
| detection | `detect_odp_unified_bytes` | odp-bytes-large | 16/16 | 0.90 | 1.52 | 1.698 | +0.62 | review trigger 5.0% | review, explained (see record) |
| detection | `detect_odp_unified_bytes` | odp-bytes-medium | 16/16 | 0.89 | 1.53 | 1.716 | +0.64 | review trigger 5.0% | review, explained (see record) |
| detection | `detect_odp_unified_bytes` | odp-bytes-tiny | 16/16 | 0.87 | 1.49 | 1.713 | +0.62 | review trigger 5.0% | review, explained (see record) |
| detection | `detect_odp_unified_reader` | odp-reader-large | 16/16 | 1.34 | 6.66 | 4.970 | +5.32 | review trigger 5.0% | review, explained (see record) |
| detection | `detect_odp_unified_reader` | odp-reader-medium | 16/16 | 1.33 | 6.68 | 5.026 | +5.36 | review trigger 5.0% | review, explained (see record) |
| detection | `detect_odp_unified_reader` | odp-reader-tiny | 16/16 | 1.28 | 6.64 | 5.167 | +5.36 | review trigger 5.0% | review, explained (see record) |

The ODS unified opens are 0.56–0.62× their leaf twins. The unified open adopts the package detection already prepared (`Spreadsheet::from_prepared_package`); the leaf twin's `Spreadsheet::from_bytes` indexes and parses on its own. That pair does not isolate the unified layer.

## Follow-up E1: the first call in a fresh process

Seven of the eight facade rows over their trigger run one timed call per
fresh child process, so one-time costs land inside the timer. In each case
the old route and the successor make the same leaf calls. In the harness,
the leaf twin of the open rows is also slower than the old route (XLSX open
352 vs 253 µs). The post-S9 harness children also take more page faults
inside the timer (median per sample: XLSX open A 99, L 132, U 144; DOCX
eager `paragraph_count` A 5, L 7, U 8.5).

Probe. `cold_open_probe.rs` makes the same calls in a small program, built at
both revisions with S1's recipe (before `d3b70a9a…`, after `4ef796df…`). Each
process opens the corpus file, the same file the harness uses (SHA-256 in
`cold-open.json`), and times the first and then a second call. It ran as 30
fresh processes per variant, pinned to CPU 20, with alternating order
(`cold_runs.py`). Values are the median of the first call in each process.

| First call in a fresh process | before | after |
|---|---:|---:|
| XLSX: facade open (`Workbook::open` / `litchi::open`) | 187.6 | 187.9 |
| XLSX: leaf open (file source + OPC index + core properties + catalog) | 182.4 (2.6 + 136.5 + 1.2 + 33.6) | 179.0 (2.4 + 136.3 + 1.2 + 34.2) |
| DOCX: facade source open (`Document::open` / `litchi::open`) | 181.1 | 180.1 |
| DOCX: leaf source open | 173.1 | 170.1 |
| DOCX: eager `document()?.paragraph_count()` | 105.1 | 104.1 |
| DOCX: eager `document()?.text()` | 75.7 | 75.8 |
| DOCX: eager separated text (what `File::text()` builds since S8a) | — | 76.4 |
| PPTX: `slide_count()` on a source-backed presentation | 0.54 | 0.47 |
| PPTX: `slide(100)` + `text_and_name()` on a source-backed presentation | 54.0 | 53.9 |

Callgrind counts the instructions of the first call as I(1 round) − I(0
rounds) of a fresh process, with the same probes:

| First call | before | after | Δ |
|---|---:|---:|---:|
| XLSX facade open | 1,130,713 | 1,139,175 | +0.75% |
| XLSX leaf open | 1,128,082 | 1,132,419 | +0.38% |
| DOCX eager text (before: `text()`; after: `text()`) | 1,249,832 | 1,251,375 | +0.12% |
| DOCX eager text (after: separated, S8a) | 1,249,832 | 1,284,430 | +2.77% |

In a fresh process of a small binary, M0a's code costs what the pre-M0a
code cost. The harness gaps come from its 265 MB executables, which differ
between the revisions (S8b added 118 routes, S9 removed 47). The probes
measure the result, not the mechanism: they show that the code costs the
same, and point to code placement and first-touch faults in the harness
binaries as the likely cause. The XLSX lifecycle's names/count/text part and
the eager DOCX open itself were not probed separately.

## Follow-up E2: allocator state in the RTF sample loop

`rtf_file_unified_open` on the 540,051-byte corpus is the eighth row (1,892 →
3,179 µs). The unified open adds only a header check, the flat-ODF probe and
one `Arc` to the shared parse. The flat-ODF probe costs 3.86 µs on this input
(`rtf-open-probe.json`). The gap is the allocator:

- `rtf_allocator_check.py` re-ran the `rtf` group twice in `A U L L U A`
  order: once with glibc's default malloc, and once with
  `GLIBC_TUNABLES=glibc.malloc.trim_threshold=4294967295:glibc.malloc.mmap_threshold=33554432`,
  which keeps freed memory mapped. There were four children per leg and
  mode, run without the GNU `time` wrapper.

  | Large corpus | default A | default L | default U | tunables A | tunables L | tunables U |
  |---|---:|---:|---:|---:|---:|---:|
  | open | 1,950 | 1,780 | 3,198 | 1,932 | 1,710 | 1,899 |
  | open + text | 1,941 | **3,193** | 1,927 | 1,904 | 1,777 | 1,890 |

  With the default malloc a ~1.3 ms slow mode appears, and which leg hits it
  depends on the process. In this re-run it was the leaf lifecycle; in the
  capture it was the unified open. With the tunables the mode is gone, and
  the unified open costs 0.98× the pre-M0a facade.
- `rtf_open_probe.rs` times unified vs leaf in one process, ABBA, 100
  samples, with the faults of the timed open:
  - bare loop (open, drop): unified 1.684–1.686 ms, leaf 1.679–1.681 ms,
    0 faults;
  - the harness's loop (untimed text and paragraph projection after the
    timer): unified 3.12–3.14 ms, leaf 2.96–2.98 ms, **2,684 faults per open**;
  - the same with the tunables: 0 faults, unified 1.85 ms, leaf 1.77 ms.

The untimed projection leaves glibc's heap such that the next timed open
re-faults about 10.5 MiB. The RTF row therefore measures the sample loop's
allocator state, not the open.

## Memory

Risk R5 says the DOC and PPT byte routes now keep their input alive. Three
measures follow.

**Allocation probe.** Neither the S8b routes (`unified_ops.rs`,
`filesystem/ops.rs`) nor the pre-M0a facade routes (`facade_ole2.rs`) open an
`allocation_metrics` region, so `litchi-perf-baseline-alloc` would report
nothing per route, and it was not built. Instead, `mem_probe.rs` wraps
`std::alloc::System` in a counting allocator, with the semantics of the
harness wrapper: requested bytes, not RSS. It opens each fixture by path and
by bytes through the revision's facade, then reads the text.

It was built twice with `mem_probe_build.sh`, using each revision's harness
`[dependencies]` and `Cargo.lock`, so leaf features and third-party versions
match the harness:

- against `a8f45144c2`, in a detached sparse worktree with its own target
  directory, removed afterwards (probe SHA-256 `be69f91c…`);
- against this branch's crates (`6bd32cb2…`).

All five measured rounds after two warm-ups were identical in every run, and
the heap was back to its starting size after every drop on both revisions.
"Held after open" is the heap still live with the handle alive. On the byte
route the window starts before `fs::read`, so retained input counts.

| Format | Fixture | Route | Input bytes | Held after open: before | after | after, `retain_source(false)` | Δ after − before | Δ / input | Peak during open: before | after |
|---|---|---|---:|---:|---:|---:|---:|---:|---:|---:|
| doc | FloatingPictures.doc | bytes | 335,360 | 711,941 | 1,061,779 | 716,709 | +349,838 | 1.04 | 1,095,905 | 1,095,905 |
| doc | FloatingPictures.doc | path | 335,360 | 711,941 | 1,061,779 | 716,709 | +349,838 | 1.04 | 1,095,905 | 1,096,017 |
| doc | documentProperties.doc | bytes | 9,728 | 22,075 | 40,985 | 26,843 | +18,910 | 1.94 | 47,709 | 47,709 |
| doc | documentProperties.doc | path | 9,728 | 22,075 | 40,985 | 26,843 | +18,910 | 1.94 | 47,709 | 47,821 |
| doc | duplicate-style-names.doc | bytes | 64,512 | refused | refused | refused | | | 92,866 | 92,866 |
| doc | duplicate-style-names.doc | path | 64,512 | refused | refused | refused | | | 92,866 | 92,978 |
| doc | ca.kwsymphony.www_education_S... | bytes | 1,619,457 | 1,090,351 | 2,731,230 | 1,095,119 | +1,640,879 | 1.01 | 3,649,201 | 3,649,201 |
| doc | ca.kwsymphony.www_education_S... | path | 1,619,457 | 1,090,351 | 2,731,230 | 1,095,119 | +1,640,879 | 1.01 | 3,649,201 | 3,649,313 |
| docx | drawing.docx | bytes | 95,247 | 120,604 | 123,500 | 123,500 | +2,896 | 0.03 | 202,517 | 202,517 |
| docx | drawing.docx | path | 95,247 | 25,389 | 28,285 | 28,285 | +2,896 | 0.03 | 107,302 | 107,302 |
| odp | tdf102223.odp | bytes | 17,335 | 88,711 | 90,207 | 90,207 | +1,496 | 0.09 | 190,198 | 190,198 |
| odp | tdf102223.odp | path | 17,335 | 88,711 | 75,468 | 75,468 | -13,243 | -0.76 | 190,198 | 158,390 |
| ods | calc-two-sheets.ods | bytes | 10,344 | 51,675 | 54,506 | 54,506 | +2,831 | 0.27 | 135,136 | 135,136 |
| ods | calc-two-sheets.ods | path | 10,344 | 42,721 | 45,552 | 45,552 | +2,831 | 0.27 | 125,850 | 125,850 |
| odt | writer-table-of-contents.odt | bytes | 44,598 | 181,383 | 182,303 | 182,303 | +920 | 0.02 | 225,733 | 225,733 |
| odt | writer-table-of-contents.odt | path | 44,598 | 139,399 | 140,319 | 140,319 | +920 | 0.02 | 180,807 | 180,807 |
| ppt | 45543-transition-litchi.ppt | bytes | 842,240 | 281,002 | 1,135,100 | 282,258 | +854,098 | 1.01 | 2,324,244 | 2,323,724 |
| ppt | 45543-transition-litchi.ppt | path | 842,240 | 255,590 | 256,846 | 256,846 | +1,256 | 0.00 | 1,434,946 | 1,434,946 |
| ppt | SampleShow.ppt | bytes | 125,440 | 76,154 | 208,248 | 77,410 | +132,094 | 1.05 | 326,418 | 325,898 |
| ppt | SampleShow.ppt | path | 125,440 | 76,154 | 83,488 | 77,410 | +7,334 | 0.06 | 200,302 | 200,302 |
| ppt | ppt_with_png.ppt | bytes | 39,424 | 36,481 | 81,607 | 37,737 | +45,126 | 1.14 | 131,622 | 131,102 |
| ppt | ppt_with_png.ppt | path | 39,424 | 45,470 | 46,726 | 46,726 | +1,256 | 0.03 | 90,593 | 90,593 |
| pptx | shapes.pptx | bytes | 68,822 | 140,505 | 141,433 | 141,433 | +928 | 0.01 | 269,526 | 269,526 |
| pptx | shapes.pptx | path | 68,822 | 71,715 | 72,643 | 72,643 | +928 | 0.01 | 200,736 | 200,736 |
| rtf | cjklist12.rtf | bytes | 44,275 | 9,787,257 | 9,791,585 | 9,791,585 | +4,328 | 0.10 | 10,202,758 | 10,202,758 |
| rtf | cjklist12.rtf | path | 44,275 | 9,787,257 | 9,791,585 | 9,791,585 | +4,328 | 0.10 | 10,202,758 | 10,202,870 |
| xlsx | Tables.xlsx | bytes | 54,387 | 116,177 | 116,782 | 116,782 | +605 | 0.01 | 245,769 | 245,769 |
| xlsx | Tables.xlsx | path | 54,387 | 61,822 | 62,427 | 62,427 | +605 | 0.01 | 191,414 | 191,414 |

- **DOC (both routes) and PPT (byte route)** keep about the input size more:
  - DOC: Δ / input 1.01–1.94. The input bytes are held plus the CFB index
    sealed into `SharedOleFile`: +1.64 MB for a 1.62 MB DOC.
  - PPT bytes: Δ / input 1.01–1.14, +854 KB for an 842 KB PPT.
  - PPT path: +1.3–7.3 KB, because that route keeps the source-backed
    package, not the bytes.
  - Peak during open is unchanged (±0.6 KB).
- **`OpenOptions::retain_source(false)`** (S5d's opt-out) removes all of it.
  DOC then holds +4,768 B and PPT +1,256 B over pre-M0a: the handle's own
  state.
- **Other formats** hold +0.6 to +4.3 KB (handle state, metadata). The
  exception is ODP by path, at −13.2 KB with 262 allocation calls instead of
  1,236: the unified path route uses the source-backed ODP owner, where the
  old facade parsed eagerly.
- `duplicate-style-names.doc` is refused by both revisions. The 1.45 MB
  `picture.doc` is too, which is why the large-DOC row uses a 1.62 MB POI
  fixture.

**Fresh-child peak RSS** (filesystem routes, per sample, from procfs in the
fresh child). This is the child's lifetime high-water mark after the
operation, so it includes the executable and the corpus preparation. Medians
over the children:

| After selector | Corpus | before route (A) MiB | unified (U) MiB | leaf (L) MiB | U − A MiB | U − L MiB |
|---|---|---:|---:|---:|---:|---:|
| `docx_file_unified_eager_full_text` | docx-source-backed-media | 40.26 | 40.86 | 40.25 | +0.60 | +0.61 |
| `docx_file_unified_eager_list_paragraphs` | docx-source-backed-media | 40.25 | 40.87 | 40.33 | +0.61 | +0.54 |
| `docx_file_unified_eager_open` | docx-source-backed-media | 40.55 | 41.15 | 40.60 | +0.60 | +0.55 |
| `docx_file_unified_eager_open_full_text_lifecycle` | docx-source-backed-media | 40.53 | 41.16 | 40.53 | +0.63 | +0.63 |
| `docx_file_unified_eager_open_paragraph_count_lifecycle` | docx-source-backed-media | 40.55 | 41.12 | 40.60 | +0.57 | +0.52 |
| `docx_file_unified_eager_paragraph_count` | docx-source-backed-media | 40.28 | 40.92 | 40.31 | +0.63 | +0.61 |
| `docx_file_unified_source_full_text` | docx-source-backed-media | 9.85 | 10.71 | 10.08 | +0.86 | +0.63 |
| `docx_file_unified_source_list_paragraphs` | docx-source-backed-media | 9.78 | 10.95 | 10.18 | +1.17 | +0.77 |
| `docx_file_unified_source_open` | docx-source-backed-media | 9.16 | 10.23 | 9.58 | +1.08 | +0.65 |
| `docx_file_unified_source_open_full_text_lifecycle` | docx-source-backed-media | 10.35 | 11.25 | 10.65 | +0.90 | +0.60 |
| `docx_file_unified_source_open_paragraph_count_lifecycle` | docx-source-backed-media | 10.29 | 11.29 | 10.79 | +1.00 | +0.50 |
| `docx_file_unified_source_paragraph_count` | docx-source-backed-media | 9.64 | 10.88 | 10.16 | +1.23 | +0.72 |
| `pptx_file_unified_eager_list_slides` | pptx-source-backed-media | 28.04 | 29.48 | 29.05 | +1.44 | +0.43 |
| `pptx_file_unified_eager_open` | pptx-source-backed-media | 26.82 | 28.47 | 28.09 | +1.66 | +0.39 |
| `pptx_file_unified_eager_open_selected_slide_lifecycle` | pptx-source-backed-media | 27.22 | 28.75 | 28.29 | +1.53 | +0.47 |
| `pptx_file_unified_eager_open_slide_count_lifecycle` | pptx-source-backed-media | 26.88 | 28.52 | 28.09 | +1.64 | +0.43 |
| `pptx_file_unified_eager_selected_slide` | pptx-source-backed-media | 27.37 | 28.76 | 28.36 | +1.39 | +0.41 |
| `pptx_file_unified_eager_slide_count` | pptx-source-backed-media | 26.88 | 28.46 | 28.00 | +1.59 | +0.46 |
| `pptx_file_unified_source_list_slides` | pptx-source-backed-media | 12.06 | 13.36 | 12.87 | +1.30 | +0.49 |
| `pptx_file_unified_source_open` | pptx-source-backed-media | 11.23 | 12.70 | 12.25 | +1.48 | +0.45 |
| `pptx_file_unified_source_open_selected_slide_lifecycle` | pptx-source-backed-media | 11.61 | 13.06 | 12.68 | +1.46 | +0.39 |
| `pptx_file_unified_source_open_slide_count_lifecycle` | pptx-source-backed-media | 11.31 | 12.75 | 12.28 | +1.44 | +0.47 |
| `pptx_file_unified_source_selected_slide` | pptx-source-backed-media | 11.27 | 12.67 | 12.13 | +1.40 | +0.54 |
| `pptx_file_unified_source_slide_count` | pptx-source-backed-media | 10.90 | 12.33 | 11.75 | +1.43 | +0.59 |
| `xlsx_file_unified_open` | xlsx-cell-values-medium | 17.00 | 10.28 | 9.67 | -6.73 | +0.61 |
| `xlsx_file_unified_open_lifecycle` | xlsx-cell-values-medium | 21.20 | 13.44 | 12.83 | -7.76 | +0.61 |
| `xlsx_file_unified_selected_cell` | xlsx-cell-values-medium | 18.69 | 11.63 | 10.97 | -7.07 | +0.66 |

U − L is +0.4 to +0.8 MiB on every filesystem route: the unified detection
and handle code touched in that process. U − A mixes binaries: DOCX +0.6 to
+1.2 MiB, PPTX +1.3 to +1.7, XLSX −6.7 to −7.8 (the pre-M0a child is the
larger one there).

**Whole-child peak RSS** (GNU `time`, per child):

| Group | A median (min–max) KiB | U median (min–max) KiB | L median (min–max) KiB |
|---|---|---|---|
| detection | 82,328 (80,788–83,036) | 199,246 (198,996–199,288) | 199,182 (199,048–199,276) |
| docx_file | 139,842 (139,708–139,976) | 199,202 (199,124–199,280) | 199,252 (199,216–199,288) |
| docx_read_one | 54,388 (53,836–55,444) | 199,206 (199,040–199,288) | 199,238 (199,184–199,284) |
| odp_file | 138,724 (136,072–140,464) | 199,220 (199,180–199,292) | 199,204 (199,156–199,288) |
| ods_file | 161,308 (159,552–161,904) | 199,204 (199,124–199,312) | 199,238 (199,020–199,268) |
| odt_file | 134,482 (133,796–137,424) | 199,232 (199,108–199,276) | 199,242 (199,176–199,280) |
| ole2_facade_large | 30,160 (30,024–30,204) | 199,234 (199,060–199,360) | 199,224 (199,000–199,300) |
| ole2_facade_refusal | 30,200 (29,932–30,228) | 199,228 (199,172–199,288) | 199,244 (199,132–199,300) |
| ole2_facade_small | 30,142 (29,880–30,200) | 199,216 (199,040–199,288) | 199,206 (199,160–199,300) |
| pptx_file | 139,718 (139,532–139,868) | 199,228 (199,184–199,280) | 199,228 (199,148–199,260) |
| rtf | 99,586 (98,904–100,392) | 199,232 (199,044–199,288) | 199,244 (199,200–199,296) |
| xlsx | 131,912 (131,056–132,124) | 199,212 (199,184–199,296) | 199,210 (199,176–199,300) |

This column cannot compare the two binaries. GNU `time` reports the largest
RSS among the child's waited-for descendants. The harness runs `git` to
record the worktree's revision and dirty state, and in this sparse worktree
that `git` process peaks at about 199 MiB. The post-S9 binary finds the
worktree and records `c475a527…, dirty`. The pre-M0a binary's lookup returns
nothing (`git_revision: null`), so its column shows the harness itself
(30–161 MiB).

S1's constant 199.4–199.7 MiB "for every group and repeat" was the same
`git` child: S1 ran the binary in its own worktree, where the lookup
succeeded. Within the post-S9 binary, U and L are equal (199.2 MiB). The
"> 1% move" tripwire S1 proposed therefore measures `git`, not the routes.
The allocation probe above is the per-route memory evidence.

## Correctness

- **Harness checks.** Every child in the tables exited 0, so every leg passed
  the harness's own untimed checks:
  - the unified, leaf and oracle projections;
  - the XLSX semantic and selected-cell digests;
  - the DOCX timed full-text oracle;
  - the OLE2 frozen outcomes;
  - the source replays.

  The failed children (§"Problems") failed on the filesystem, not on a check.
- **Outputs across legs** (`outputs` in `analysis.json`):
  - **XLSX.** The semantic digest `020fdd14…` and the source digest
    `dfff7ec0…` are identical on all three legs.
  - **DOC and PPT.** The old facade's frozen observation (for example
    `text:9017:c924…` for `FloatingPictures.doc`) equals the unified one on
    all 12 successful rows (two DOC and two PPT fixtures, three scenarios
    each).
  - **Refusal.** All three legs refuse `duplicate-style-names.doc`. The old
    facade and the leaf report the leaf message ("Corrupted file: invalid
    stylesheet: style names and aliases must be unique"). The unified leg
    reports its code-level message ("corrupt_file: the file is malformed
    [Doc]") and keeps the leaf in `Error::leaf()`, as S8b recorded. These
    are the only rows whose observations differ.
  - **DOCX full text.** 10,000 bytes before and 10,199 after: one `\n`
    between the 200 paragraphs, the S8a B3 change. With the harness fix,
    unified and leaf time the same 10,199-byte text (digest `b4b33b96…`).
  - **PPT and DOC text.** The placeholder, table and quote fixes are only on
    `uo/m1`, so no difference was expected here, and none was seen.

## Environment: load per run

The 1-minute load average before each child ranged from 2.55 to 21.12 on 32
CPUs, as recorded before or after any child. Eleven children waited at the
load gate, 600 s in total, against the 2-hour cap. Per-child values are in
`children.jsonl`.

| Group | Block | Children | Start (UTC) | Minutes | 1-min load before each child (min–max) | Gate waits |
|---|---|---|---|---|---|---|
| docx_read_one | 1 | 6 | 2026-10-05T18:48:51 | 0.3 | 2.86–2.96 | 0 s |
| odt_file | 1 | 6 | 2026-10-05T18:49:08 | 0.3 | 2.68–2.96 | 0 s |
| rtf | 1 | 6 | 2026-10-05T18:49:28 | 0.2 | 2.63–2.82 | 0 s |
| ole2_facade_small | 1 | 6 | 2026-10-05T18:49:38 | 0.0 | 2.75–2.82 | 0 s |
| ole2_facade_large | 1 | 6 | 2026-10-05T18:49:40 | 0.0 | 2.75–2.75 | 0 s |
| ole2_facade_refusal | 1 | 6 | 2026-10-05T18:49:43 | 0.0 | 2.75–2.77 | 0 s |
| detection | 1 | 6 | 2026-10-05T18:49:44 | 0.1 | 2.77–4.55 | 0 s |
| docx_read_one | 2 | 6 | 2026-10-05T18:49:53 | 0.3 | 4.55–7.02 | 0 s |
| odt_file | 2 | 6 | 2026-10-05T18:50:10 | 0.3 | 5.29–7.82 | 30 s |
| rtf | 2 | 6 | 2026-10-05T18:50:59 | 0.2 | 4.86–5.11 | 0 s |
| ole2_facade_small | 2 | 6 | 2026-10-05T18:51:10 | 0.0 | 4.63–4.63 | 0 s |
| ole2_facade_large | 2 | 6 | 2026-10-05T18:51:12 | 0.0 | 4.50–4.63 | 0 s |
| ole2_facade_refusal | 2 | 6 | 2026-10-05T18:51:14 | 0.0 | 4.50–4.50 | 0 s |
| detection | 2 | 6 | 2026-10-05T18:51:16 | 0.1 | 4.46–4.50 | 0 s |
| docx_read_one | 3 | 6 | 2026-10-05T18:51:25 | 0.3 | 3.99–4.26 | 0 s |
| odt_file | 3 | 6 | 2026-10-05T18:51:41 | 0.3 | 3.42–3.83 | 0 s |
| rtf | 3 | 6 | 2026-10-05T18:52:01 | 0.2 | 3.20–3.31 | 0 s |
| ole2_facade_small | 3 | 6 | 2026-10-05T18:52:12 | 0.0 | 3.27–3.27 | 0 s |
| ole2_facade_large | 3 | 6 | 2026-10-05T18:52:14 | 0.0 | 3.27–3.33 | 0 s |
| ole2_facade_refusal | 3 | 6 | 2026-10-05T18:52:16 | 0.0 | 3.33–3.33 | 0 s |
| detection | 3 | 6 | 2026-10-05T18:52:18 | 0.1 | 3.20–3.33 | 0 s |
| docx_read_one | 4 | 6 | 2026-10-05T18:52:27 | 0.3 | 3.18–4.84 | 0 s |
| odt_file | 4 | 6 | 2026-10-05T18:52:43 | 0.3 | 4.77–5.20 | 0 s |
| rtf | 4 | 6 | 2026-10-05T18:53:03 | 0.2 | 4.94–5.20 | 0 s |
| ole2_facade_small | 4 | 6 | 2026-10-05T18:53:14 | 0.0 | 4.86–4.94 | 0 s |
| ole2_facade_large | 4 | 6 | 2026-10-05T18:53:16 | 0.0 | 4.86–4.86 | 0 s |
| ole2_facade_refusal | 4 | 6 | 2026-10-05T18:53:19 | 0.0 | 4.71–4.86 | 0 s |
| detection | 4 | 6 | 2026-10-05T18:53:20 | 0.1 | 4.71–5.22 | 0 s |
| docx_read_one | 5 | 6 | 2026-10-05T18:53:29 | 0.3 | 5.22–5.96 | 0 s |
| odt_file | 5 | 6 | 2026-10-05T18:53:45 | 0.3 | 5.65–6.00 | 0 s |
| rtf | 5 | 6 | 2026-10-05T18:54:05 | 0.2 | 5.28–5.69 | 0 s |
| ole2_facade_small | 5 | 6 | 2026-10-05T18:54:16 | 0.0 | 5.28–5.28 | 0 s |
| ole2_facade_large | 5 | 6 | 2026-10-05T18:54:18 | 0.0 | 5.26–5.28 | 0 s |
| ole2_facade_refusal | 5 | 6 | 2026-10-05T18:54:21 | 0.0 | 5.26–5.26 | 0 s |
| detection | 5 | 6 | 2026-10-05T18:54:23 | 0.1 | 5.26–5.67 | 0 s |
| docx_read_one | 6 | 6 | 2026-10-05T18:54:32 | 0.3 | 5.44–5.70 | 0 s |
| odt_file | 6 | 6 | 2026-10-05T18:54:48 | 0.3 | 4.94–5.44 | 0 s |
| rtf | 6 | 6 | 2026-10-05T18:55:08 | 0.2 | 4.94–5.27 | 0 s |
| ole2_facade_small | 6 | 6 | 2026-10-05T18:55:20 | 0.0 | 5.49–5.49 | 0 s |
| ole2_facade_large | 6 | 6 | 2026-10-05T18:55:22 | 0.0 | 5.49–5.49 | 0 s |
| ole2_facade_refusal | 6 | 6 | 2026-10-05T18:55:24 | 0.0 | 5.37–5.37 | 0 s |
| detection | 6 | 6 | 2026-10-05T18:55:26 | 0.2 | 5.37–7.52 | 120 s |
| docx_read_one | 7 | 6 | 2026-10-05T18:57:37 | 0.3 | 6.82–7.52 | 0 s |
| odt_file | 7 | 6 | 2026-10-05T18:57:53 | 0.3 | 6.30–6.82 | 0 s |
| rtf | 7 | 6 | 2026-10-05T18:58:13 | 0.2 | 6.03–6.30 | 0 s |
| ole2_facade_small | 7 | 6 | 2026-10-05T18:58:24 | 0.0 | 5.94–6.03 | 0 s |
| ole2_facade_large | 7 | 6 | 2026-10-05T18:58:26 | 0.0 | 5.94–5.94 | 0 s |
| ole2_facade_refusal | 7 | 6 | 2026-10-05T18:58:29 | 0.0 | 5.79–5.94 | 0 s |
| detection | 7 | 6 | 2026-10-05T18:58:31 | 0.1 | 5.64–5.79 | 0 s |
| docx_read_one | 8 | 6 | 2026-10-05T18:58:39 | 0.3 | 5.28–5.64 | 0 s |
| odt_file | 8 | 6 | 2026-10-05T18:58:55 | 0.3 | 4.62–5.10 | 0 s |
| rtf | 8 | 6 | 2026-10-05T18:59:15 | 0.2 | 4.49–5.99 | 0 s |
| ole2_facade_small | 8 | 6 | 2026-10-05T18:59:27 | 0.0 | 5.99–5.99 | 0 s |
| ole2_facade_large | 8 | 6 | 2026-10-05T18:59:28 | 0.0 | 5.75–5.99 | 0 s |
| ole2_facade_refusal | 8 | 6 | 2026-10-05T18:59:31 | 0.0 | 5.75–5.75 | 0 s |
| detection | 8 | 6 | 2026-10-05T18:59:33 | 0.1 | 5.33–5.75 | 0 s |
| odp_file | 1 | 6 | 2026-10-05T18:59:42 | 0.8 | 5.33–6.59 | 0 s |
| ods_file | 1 | 6 | 2026-10-05T19:00:32 | 1.7 | 4.43–5.54 | 0 s |
| odp_file | 2 | 6 | 2026-10-05T19:02:16 | 0.8 | 3.82–4.15 | 0 s |
| ods_file | 2 | 6 | 2026-10-05T19:03:05 | 1.7 | 3.16–5.13 | 0 s |
| odp_file | 3 | 6 | 2026-10-05T19:04:48 | 1.3 | 3.24–4.62 | 0 s |
| ods_file | 3 | 6 | 2026-10-05T19:06:09 | 3.5 | 2.77–4.56 | 0 s |
| odp_file | 4 | 6 | 2026-10-05T19:09:40 | 1.5 | 2.88–4.97 | 0 s |
| ods_file | 4 | 6 | 2026-10-05T19:11:10 | 2.8 | 5.27–6.60 | 0 s |
| xlsx | 1 | 6 | 2026-10-05T19:13:57 | 6.8 | 4.31–7.12 | 30 s |
| xlsx | 2 | 6 | 2026-10-05T19:21:16 | 6.7 | 5.26–7.69 | 0 s |
| xlsx | 3 | 6 | 2026-10-05T19:28:00 | 6.8 | 5.84–7.92 | 90 s |
| pptx_file | 4 | 6 | 2026-10-05T20:54:38 | 33.9 | 2.55–6.22 | 180 s |
| pptx_file | 5 | 6 | 2026-10-05T21:32:04 | 34.0 | 4.29–7.85 | 120 s |
| docx_file | 1 | 6 | 2026-10-05T22:08:02 | 50.3 | 3.66–7.24 | 30 s |

## Problems and deviations

1. **Harness defect at `b906691f3b`** (§"Harness fix"). Every DOCX filesystem
   route failed its own verification. This was fixed in `d10dd4c747` before
   the measured build.
2. **Shared tmpfs quota.** `/tmp` is a tmpfs with a per-user quota of 49.38
   GiB. Every agent on the host runs as the same user, and other agents'
   build directories held almost all of it. The harness writes one corpus
   copy per filesystem case into the temporary directory, 12 × 17 MB per PPTX
   child. Four PPTX children failed with `Disk quota exceeded`:
   - `B1-2U` and `B2-1A`;
   - `B3-2U` and `B3-3L`, six attempts each under the retry added for this.

   The driver also lost one `children.jsonl` line to the quota.
3. **My error.** `B1-4L-pptx_file` failed because I replaced the after
   binary's scratch copy with a symbolic link while that child's parent was
   running, and the harness spawns its fresh children through
   `current_exe()`. Children started after 19:49:43 UTC ran the same bytes
   from `target/release`. Before that, the in-process groups and XLSX ran a
   byte-identical copy on `/tmp`.
4. **Interruption and re-run.** The session was stopped at about 20:07 UTC,
   which ended the main capture chain before DOCX.
   - The partial tmpfs PPTX blocks 1–3 are set aside: they are not in the
     tables but are in the raw tarball.
   - PPTX blocks 4–5 and DOCX block 1 were re-run with `--filesystem-root
     <worktree>/target/s10-fsroot` (ext4) for all three legs, because the
     tmpfs could not hold the corpus copies.
   - Absolute PPTX times differ between the two filesystems (eager open
     7.8 ms on tmpfs, 7.0 ms on ext4), so the PPTX and DOCX rows use only
     the ext4 blocks. Every other group ran on tmpfs, as in S1.
5. **Fewer children for the slow groups.** PPTX has 2 blocks (4 children per
   leg) and DOCX 1 block (2 per leg). The rule's minimum of two children per
   leg holds. DOCX's S1 noise is at most 3.4%.
6. **Raw data.** The raw reports are not committed (size). The pruned bundles
   keep every field the analysis reads, and the analysis reproduces from
   them exactly (checked on the in-process groups). The full set is cited
   under §"Files".
7. **No allocator binary.** It was not built, because it would not report
   these routes. The allocation probe replaces it (§"Memory").
8. **Not gateable as is.** The whole-child RSS tripwire (§"Memory"), and
   S1's spreads for rows under 5 µs, which the floor and separation rules
   handle.
9. **Not committed.** The probe and harness binaries (AGENTS.md); their
   SHA-256 values are in `environment.json`.
