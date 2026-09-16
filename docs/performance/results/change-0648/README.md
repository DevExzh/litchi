# Evidence packet — change 0648

The XLS shared-string resolver retains the workbook's string table for the life
of one scan, once that scan has shown it is walking the table, and decodes each
entry from a slice of it. Record:
[`docs/performance/0648-xls-shared-string-resolver-window.md`](../../0648-xls-shared-string-resolver-window.md).

## Provenance

| | |
| --- | --- |
| base commit | `9f28ea621dcb3c0ba6b19f1e7cf1c1f6b3d0e5f5` (`feat/office-format-completeness`, change 0636) |
| branch | `perf/0648-xls-shared-string-resolver-window` |
| commit | the tip of that branch — a commit cannot record its own hash, so neither this table nor `decision.json` names one |
| before checkout | `/home/zhuhe/code/litchi-worktrees/before-9f28ea621`, a read-only detached checkout of the base |
| host | AMD EPYC, 32 cores, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0, cargo 1.95.0, valgrind 3.26.0, linux-perf |
| build | `cargo build --release --locked`; workspace `[profile.release] lto = true, panic = "abort"` |
| pinning | `taskset -c 23` for every measured process |
| host load | seven other agents building and measuring throughout; load average 13–31 across the windows. Every timing window carries its own A/A floor, and the floor is under 1.4% at p50 on every scenario |
| binaries | `binaries.sha256` — the two attribution probes, the two `litchi-perf-baseline` builds and the two `xls_source_attribution` builds, all staged outside their Cargo target directories before any leg ran. A final defensive edit changed the after probe after the first measurement window, so every count, pair and timing leg was re-taken against the binaries named there; the first windows are kept in `cycles/cycles-run-history.txt` and `timing/timing-run-history.txt` |

## Contents

| path | what it is |
| --- | --- |
| `binaries.sha256` | SHA-256 of the four staged binaries |
| `probe/README.md`, `probe/main.rs`, `probe/Cargo.toml.template` | the attribution probe, its five modes and how to build it against each leg |
| `counts/xls-fixtures.txt` | the 126 `.xls` files under `test-data/` the sweep visits |
| `counts/before-sweep.tsv`, `counts/after-sweep.tsv` | the sweep verbatim: 565 rows of `path, operation, reads, bytes, reads ≤ 512 B, observations, len calls, outcome digest` |
| `counts/sweep-summary.txt` | the per-operation totals the record cites, and the largest per-fixture byte and read movements in both directions |
| `counts/alloc.tsv` | allocations and allocated bytes for `open`, `all-cells` and `full-text` on three fixtures, both legs, from the probe's own counting allocator |
| `counts/attribution-54016.txt` | change 0605's `xls_source_attribution` binary on `54016.xls` in `owned-readat` and `file-source` modes for `all-cells` and `full-text`, both legs: an independent instrument with its own counters and its own oracle |
| `counts/attrib-corpus.sh`, `counts/attrib-corpus-before.tsv`, `counts/attrib-corpus-after.tsv`, `counts/attrib-corpus-summary.txt` | the same binary swept over the whole corpus, 252 rows per leg, with its own outcome string on every row: a second differential through a second instrument |
| `counts/sst-region-sizes.tsv` | every fixture's SST record-group extent, walked out of its Workbook stream: 108 fixtures with a table, 61 distinct sizes, 8 bytes to 225,003, none above the 256 KiB ceiling |
| `trace/*.tsv.gz` | one line per resolve — the SST region's bounds, and the entry's source offset and span — for four fixture/operation pairs |
| `trace/instrumentation.patch.md` | the temporary `eprintln!` that produced them, and how to regenerate them |
| `trace/simulate.py`, `trace/policy-simulation.txt` | the policy simulator and its output: six sliding-window ceilings with and without change 0568's growth schedule, a span-sized window, and the whole-table policy. This is the evidence for rejecting the sliding window |
| `callgrind/callgrind.sh`, `callgrind/cg-*.log`, `callgrind/callgrind-summary.txt` | callgrind isolation pairs N=2 and N=6, differenced and divided by 4 |
| `cycles/cycles.sh`, `cycles/perf-*.txt`, `cycles/perfsum.py`, `cycles/cycles-summary.txt` | `perf stat -r 5` isolation pairs at N=20 and N=120, differenced and divided by 100 |
| `cycles/cycles-run-history.txt` | both `perf stat` windows this change ran, cycles beside instructions: why the record reads effect sizes from the instruction column |
| `timing/time.sh`, `timing/t-*.txt`, `timing/stats.py`, `timing/timing-summary.txt`, `timing/timing-table.txt` | the paired runner, every leg verbatim at 120 samples, and the percentile and paired-delta tables with the A/A and B/B floors |
| `timing/timing-run-history.txt` | both paired-timing windows this change ran, and the tail in the second window's A legs that its p95 and p99 read |
| `range/before-range-54016.json.gz`, `range/after-range-54016.json.gz` | change 0627's five XLS range-source selectors on `54016.xls`, both legs, schema-1 reports carrying the per-sample physical request counts, bytes and request-sequence digests |
| `range/range-compare.txt` | the comparison: requests, bytes, digest identity, frozen-observation identity and the modelled service floor on both legs |
| `gates.txt` | the tail and exit code of every gate |
| `decision.json` | the decision record |
| `log-sections.md` | the four log paragraphs the coordinator merges |

## Replaying the counts

```sh
# for each leg: write probe/Cargo.toml from the template with __TREE__ replaced
# by that leg's checkout, then
CARGO_TARGET_DIR=<scratch> cargo build --release
taskset -c 23 <binary> sweep $(cat counts/xls-fixtures.txt | tr '\n' ' ')
```

The two TSV files must agree on the outcome digest of all 565 rows, and column
for column on every `open`, `list` and `validate` row.

```sh
taskset -c 23 <binary> alloc all-cells test-data/poi/test-data/spreadsheet/54016.xls
taskset -c 23 <binary> sst  $(cat counts/xls-fixtures.txt | tr '\n' ' ')
```

## Replaying the range-source legs

```sh
taskset -c 23 litchi-perf-baseline --warmup 1 --samples 3 \
  --case xls_range_source_open,xls_range_source_open_list_worksheets,\
xls_range_source_open_one_cell,xls_range_source_open_all_cells,\
xls_range_source_open_full_text \
  --ole2-file test-data/poi/test-data/spreadsheet/54016.xls \
  --range-fixed-latency-us 1000 --range-request-overhead-us 0 \
  --range-bandwidth-bytes-per-sec 104857600 --range-max-physical-bytes 65536 \
  --json <out>.json
```

Three retained samples rather than change 0627's fifty: the quantities compared
here — physical request counts, bytes and the request-sequence digest — are
deterministic and the selector's own identity gate fails if they are not, while
the elapsed figure is sleep-driven model arithmetic that this change does not
claim. The before leg reproduces 0627's published counts exactly (40, 40, 66,
16,145, 16,145), which is the provenance check on this packet's baseline. The
before leg's `all-cells` and `full-text` cases take about 16 s per sample by
construction, which is the result.
