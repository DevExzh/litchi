# Evidence packet — change 0636

A bounded window over `SharedOleStreamCursor`, added as the wrapper type
`litchi_cfb::BufferedOleStreamCursor` and enabled at one caller, the XLS
validation record walk. Record:
[`docs/performance/0636-cfb-cursor-bounded-window.md`](../../0636-cfb-cursor-bounded-window.md).

## Provenance

| | |
| --- | --- |
| base commit | `c7326f68065edf6f2198ca3cb39c38c48cf00ed9` (`feat/office-format-completeness`, change 0630) |
| branch | `perf/0636-cfb-cursor-range-read-ahead` |
| commit | see `decision.json` |
| before checkout | `/home/zhuhe/code/litchi-worktrees/before-c7326f680`, a read-only detached checkout of the base |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0, cargo 1.95.0, valgrind 3.26.0, linux-perf |
| build | `cargo build --release --locked`; workspace `[profile.release] lto = true, panic = "abort"` |
| pinning | `taskset -c 12` for every measured process |
| host load | seven other agents building and measuring throughout; load average 12–46 across the windows. Every timing window carries its own A/A floor, and the two scenarios whose sign was unstable were re-run twice more |
| binaries | `binaries.sha256` — the two attribution probes and the two `litchi-perf-baseline` builds |

The probe binaries were copied out of their Cargo target directories before any
leg ran, per change 0627's finding that a concurrent build can relink a binary
mid-run.

## Contents

| path | what it is |
| --- | --- |
| `binaries.sha256` | SHA-256 of the four staged binaries |
| `probe/main.rs` | the attribution probe: counts logical positional reads, read bytes and source-version observations taken by the public `litchi_xls` APIs over one in-process `ReadAt`, and freezes each operation's outcome as a digest so a before/after pair is also a differential. Modes: `sweep` (one TSV line per fixture and operation), `detail`, `time <owned\|file> <op> <n> <file>` |
| `probe/Cargo.toml.template` | its manifest, with the tree path parameterised as `__TREE__`; build it against each leg's checkout with its own `CARGO_TARGET_DIR` |
| `counts/xls-fixtures.txt` | the 126 `.xls` files under `test-data/` the sweep visits |
| `counts/before-sweep.tsv`, `counts/after-sweep.tsv` | the sweep verbatim: 565 rows of `path, operation, reads, bytes, reads ≤ 512 B, observations, len calls, outcome digest` |
| `counts/sweep-summary.txt` | the per-operation totals the record's counts table cites, the maximum per-fixture over-read, and the differential verdict — the awk that produced it fails loudly on any outcome-digest difference or any change to a read-path row |
| `cycles/cycles.sh` | the `perf stat` runner: N=20 and N=120 iterations, `-r 5`, differenced and divided by 100 |
| `cycles/perf-*.txt` | its raw CSV output, 28 files |
| `cycles/cycles-run-history.txt` | the read-path control deltas across all three `perf stat` windows this change ran, cycles beside instructions: the reason the record reads those rows from instructions and paired medians |
| `cycles/perfsum.py`, `cycles/cycles-summary.txt` | the differencing script and the table the record cites |
| `callgrind/callgrind.sh` | the callgrind runner: isolation pairs N=2 and N=6 |
| `callgrind/*.log` | the summary of each callgrind run, carrying the `I refs` total |
| `callgrind/callgrind-summary.txt` | the differenced pairs, with the `rep movsb` caveat that makes one row disagree with `perf stat` |
| `timing/time.sh` | the paired runner: order A1 B1 B2 A2, one pinned CPU, warm-up discarded by the shell |
| `timing/t-<scenario>-<leg>.txt` | every leg verbatim, one elapsed nanosecond figure per sample, 120 samples per leg |
| `timing/stats.py`, `timing/timing-summary.txt` | the percentile and paired-delta script, and its output |
| `timing/timing-table.txt` | the same legs as one table with the A/A and B/B floors beside each paired delta |
| `timing/repeats/` | two further independent runs at 150 samples of the two scenarios whose sign was unstable (`full-text` and `all-cells` on `54016.xls`), with their own summary |
| `range/before-range-54016.json.gz`, `range/after-range-54016.json.gz` | change 0627's five XLS range-source selectors on `54016.xls`, both legs, schema-1 reports carrying the per-sample physical request counts, bytes and request-sequence digests |
| `range/range-compare.txt` | the comparison: requests, bytes and the complete ordered request-sequence SHA-256 on both legs |
| `gates.txt` | the tail and exit code of every gate |
| `decision.json` | the decision record |
| `log-sections.md` | the four log paragraphs the coordinator merges |

## Replaying the counts

```sh
# for each leg: write probe/Cargo.toml from the template with __TREE__ replaced
# by that leg's checkout, then
CARGO_TARGET_DIR=<scratch> cargo build --release
taskset -c 12 <binary> sweep $(cat counts/xls-fixtures.txt | tr '\n' ' ')
```

The two TSV files must agree column for column on every `open`, `list`,
`all-cells` and `full-text` row, and on the outcome digest of all 565 rows.

## Replaying the range-source legs

```sh
taskset -c 12 litchi-perf-baseline --warmup 1 --samples 3 \
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
16,145, 16,145), which is the provenance check on this packet's baseline.
