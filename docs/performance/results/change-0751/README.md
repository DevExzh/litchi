# Change 0751 evidence packet

Record: [`../../0751-pptx-cross-copy-apply-digest-reuse.md`](../../0751-pptx-cross-copy-apply-digest-reuse.md).
Base `6d989cad63`. Production commits `4c7cbc4b8c` (`litchi-opc`) and
`f84522f8af` (`litchi-pptx`). After the independent review (verdict: merge
after fixes), `aa606e0016` (facade memo re-projection), `c0fefbc85b` (the
`litchi-opc` nits) and `098f6ffd4e` (extended golden transcript). Every table
of the record is from the matrix of `098f6ffd4e`.

## Contents

| path | what it is |
|---|---|
| `raw/native/<case>/r<round>-s<slot>-<arm>.json.gz` | every native ABBA harness report of the matrix of `098f6ffd4e`, 4 rounds × (before, after, after, before) per case, core 4 |
| `raw/native/<case>/…perf.csv.gz` | the whole-child `perf stat -e cycles,instructions` of that process |
| `raw/alloc/…` | the allocator lane, same layout |
| `raw/*/receipts.jsonl.gz`, `raw/*.log` | argv, binary SHA-256, start/end times and exit code of every process, and the runners' progress logs |
| `counters/isolation-pairs/` | `perf stat` isolation pairs (`--samples 12` minus `--samples 2`, ÷ 10) for the media-rich and plain lifecycles and the semantic control; `isolation-pairs.json.gz` is the summary |
| `counters/tiny-semantic/` | the dedicated tiny-corpus semantic control runs: isolation pairs at 1,100 and 100 samples in two ABBA rounds, and whole-child front-end counters (`fe-*.csv.gz`) |
| `analysis/matrix.json`, `analysis/matrix.md` | `scripts/analyze.py raw`: per-process rows, per-arm medians, paired ratios with bootstrap intervals, phases, counters, allocator fields |
| `analysis/faults.json` | `scripts/faults.py raw`: media-rich lifecycle samples grouped by fresh 33.6 MB mappings |
| `analysis/tiny-semantic.json` | `scripts/tiny_semantic.py counters/tiny-semantic` |
| `attribution/before-attribution.json`, `attribution/after-attribution.json` | `scripts/attribute.py` over 12-sample `perf record -e cycles --call-graph fp` captures of the frame-pointer builds of `6d989cad63` and `f84522f8af` (the raw `perf.data` is not retained, 53 MB) |
| `golden/base-6d989cad63-transcript.txt`, `golden/after-transcript.txt` | the 196-line transcript `digest_reuse_tests.rs` records, printed on the base tree and on this branch; both SHA-256 `32b7dbf98c16d2a9d0a1151ecc4a4b10f008ca3d2b638696e3c5abc720c910b0`. Its first 98 lines are the first version's transcript (`cb410eef…`) |
| `binaries.sha256` | the measured binaries of `098f6ffd4e` and the rebuilt base |
| `gates.txt` | gate commands, exit codes and test counts after the review |
| `superseded-f84522f8af/` | the first matrix, of `f84522f8af` before the review: its analysis (every process's p50, p95, mean, counters and allocator medians), fault regrouping, tiny-control summary, binaries, gates, receipts and cleanup record. Its raw reports were replaced by this matrix's |
| `cleanup.json` | what was removed and kept |
| `log-sections.md` | paragraphs for `HOTSPOTS.md`, `REPORT.md` and `GOAL_AUDIT.md` |
| `scripts/` | `run_abba.py`, `analyze.py`, `isolation_pairs.py`, `tiny_semantic.py`, `faults.py` (0742's, unchanged apart from reading gzip), `attribute.py`, `gates.sh`, `cleanup.py` |

## Reproduce the tables

```sh
cd docs/performance/results/change-0751
python3 scripts/analyze.py raw --json analysis/matrix.json --markdown analysis/matrix.md
python3 scripts/faults.py raw --json analysis/faults.json
python3 scripts/tiny_semantic.py counters/tiny-semantic --json analysis/tiny-semantic.json
```

## How the runs were made

```sh
# native lane (every process under perf stat)
run_abba.py --lane native --before before-native --after after-native --core 4 --rounds 4 \
    --samples 20 --warmup 3 --perf-stat pptx_cross_copy_media_rich_lifecycle pptx_cross_copy_media_rich
run_abba.py --lane native ... --samples 40 --warmup 3 --perf-stat \
    pptx_cross_copy_plain_lifecycle pptx_cross_copy_plain pptx_source_backed_cross_copy_media_rich_lifecycle
run_abba.py --lane native ... --samples 200 --warmup 20 --perf-stat pptx_semantic_one_edit_save
# allocator lane
run_abba.py --lane alloc --before before-alloc --after after-alloc --core 4 --rounds 4 \
    --samples 3 --warmup 1 pptx_cross_copy_media_rich_lifecycle pptx_cross_copy_plain_lifecycle
# per-iteration counters
isolation_pairs.py --before before-native --after after-native --core 4 --low 2 --high 12 \
    --warmup 1 --repeats 2 --out isolation \
    pptx_cross_copy_media_rich_lifecycle pptx_cross_copy_plain_lifecycle pptx_semantic_one_edit_save
# tiny semantic control: see scripts/tiny_semantic.py's docstring
```

Both legs were built with the identical command,
`cargo build --release --locked --offline --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline`
(plus `--features allocator-metrics --bin litchi-perf-baseline-alloc`), the
before leg from a detached worktree of `6d989cad63`, the after leg from
`098f6ffd4e`, and staged outside the target directories before any run. The
frame-pointer builds added `RUSTFLAGS="-C force-frame-pointers=yes"` and were
used for attribution only.

`performance_claim: none`.
