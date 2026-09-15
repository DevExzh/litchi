# Retained evidence — change 0624

Item **CORE-4 / SAVE-6** of change 0587: the gate measurement for per-member
parallel deflate of an OOXML save's changed member set, and the frozen design it
justifies. Record:
[`docs/performance/0624-parallel-changed-member-deflate-design.md`](../../0624-parallel-changed-member-deflate-design.md).

**No production file was modified by this change.** Everything here is
measurement plus two retained scratch probes.

## Provenance

| | |
| --- | --- |
| base commit | `aba3f943d` (change 0618, `perf(zip): drive one Deflate compressor per authored save, not one per member`) |
| branch | `perf/0624-parallel-changed-member-deflate-design` |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| toolchain | `rustc 1.95.0 (59807616e 2026-04-14)`, `perf` 7.0.14 |
| workspace codec versions | `flate2` 1.1.10, `zlib-rs` 0.6.7, `rayon` 1.12.0, `rayon-core` 1.13.0 (workspace `Cargo.lock`) |
| pinning | deterministic counts `taskset -c 18`; scaling legs `taskset -c 18-25` |
| host state | seven other agents active throughout; run-window load average 7.7-20.5 |
| `probe0624` sha256 | `69516155c58b2f9dc04bcad222fd808eaec890298f32262e243b1ec322f37244` |
| `probe0624-deflate` sha256 | `dc35e7b44bb6ff5fd1d82955044ad9719b0372488776fc82c9ad98257f23380a` |

The two sha256s above are of binaries built from the **formatted** sources
retained here, after `cargo fmt` was run on both probes at the end of the batch.
The measurements in the record were taken with binaries built from the same
sources before formatting; the reformatted binaries were re-run as a spot check
and reproduce the census byte-identically (`xlsxspread:1` on the dense-wide
shape: 7 members, 2 changed, 1,893,450 B each) and the scaling within its stated
floor (40 × ~871 B at width 4: 3.534× against the 3.671× the record cites, with
a width-2 A/A floor of 3.62%).

Both probes were built `--release --offline` with `lto = true`, each with its own
`CARGO_TARGET_DIR` outside the checkout. Neither uses `--locked`: each is a fresh
crate with no lockfile of its own. `probe0624-deflate` has **no litchi
dependency** on purpose, and its `zlib-rs` was downgraded to the workspace's
0.6.7 with `cargo update -p zlib-rs --precise 0.6.7 --offline` so its codec is
bit-for-bit the one the writers run.

## Contents

| path | what it is |
| --- | --- |
| `probe/probe0624/src/main.rs` | the save-census and save-split probe: drives the `litchi-xlsx` value editor, the `OpcPackage` publication route and the `litchi-pptx` authored route; diffs published against source member by member; rebuilds the `tools/perf-baseline` XLSX corpus shapes; times a save against the deflate of its own changed set |
| `probe/probe0624/Cargo.toml.example` | manifest template for the above |
| `probe/probe0624-deflate/src/main.rs` | the deflate-scaling probe: the writers' codec, serial against a caller-owned Rayon pool, over member sets of a stated count and size |
| `probe/probe0624-deflate/Cargo.toml.example` | manifest template, including the `zlib-rs` pin |
| `census/census-opc.tsv` | 325 fixtures × 2 `OpcPackage` scenarios, 672 rows: members, changed members, changed bytes, largest changed member, changed members ≥256 KiB |
| `census/census-xlsx.tsv` | 157 workbooks × 2 `litchi-xlsx` editor scenarios, 414 rows, same columns; includes the 66 non-XLSX and 38 editor refusals |
| `census/authored-census.txt` | authored PPTX decks at 1, 50 and 200 slides: members, total uncompressed bytes, largest member, members ≥256 KiB |
| `timing/perf-dense-spread.txt` | `perf report` for the harness dense-wide one-percent save, with the summed zlib/deflate cycle share |
| `timing/perf-ndp-onecell.txt` | the same for `no_drawing_patriarch.xlsx`, one cell |
| `timing/perf-pptx50.txt` | the same for the authored 50-slide PPTX save |
| `timing/perf-cfs-addrelall.txt` | the same for `ConditionalFormattingSamples.xlsx`, a relationship on every related part (40 regenerated members) |
| `timing/perf-cfs-onecell.txt` | the same for `ConditionalFormattingSamples.xlsx`, one cell |
| `timing/savesplit.txt` | publish p50/mean/p95/p99 against deflate-only p50/mean/p95/p99 for eight scenarios, 200-500 samples per leg |
| `scaling/scale-real-sets.txt` | five of the seven real changed sets at widths 1/2/4/8: p50/mean/p95/p99 per leg, speedup, efficiency, Amdahl serial fraction, cell class |
| `scaling/scale-authored.txt` | the same for the two authored-PPTX member sets (161 and 537 members), dumped from the published decks |
| `scaling/aa-floor.txt` | four repeats of the serial and width-2 legs on three sets, same window — the A/A floor |
| `scaling/crossover.txt` | the controlled crossover sweep: 871-byte members, counts 4 to 256, widths 1/2/4 |
| `scaling/sweep-xml.txt` | the wide synthetic sweep: 7 member sizes × 6 counts × 4 widths |
| `scaling/sweep-small.txt` | the small-payload re-run whose cells did **not** reproduce; retained as the reason the sweep's smallest and width-8 cells are reported out-of-model |
| `gates.txt` | the tail of every gate run |
| `decision.json` | the decision record |
| `log-sections.md` | the four log paragraphs for HOTSPOTS, GOAL_AUDIT, REPORT and ADR_COMPLIANCE (the coordinator merges these) |

## How to reproduce

```sh
# both probes, each with its own target directory, from the checkout root
cp docs/performance/results/change-0624/probe/probe0624/Cargo.toml.example probe0624/Cargo.toml
sed -i "s|REPLACE_WITH_CHECKOUT|$PWD|g" probe0624/Cargo.toml
cp docs/performance/results/change-0624/probe/probe0624/src/main.rs probe0624/src/main.rs
CARGO_TARGET_DIR=/tmp/none cargo build --release --offline --manifest-path probe0624/Cargo.toml

# the census (deterministic; no timing sensitivity)
taskset -c 18 probe0624 census test-data addrel addrelall
taskset -c 18 probe0624 census test-data/ooxml/xlsx xlsxcells:1 xlsxpercent:1

# the harness corpus shapes, then their changed sets
taskset -c 18 probe0624 synthwb 2 256 256 dense-wide.xlsx
taskset -c 18 probe0624 censusone dense-wide.xlsx xlsxspread:1

# deflate's share of a save, two ways
taskset -c 18 probe0624 savesplit dense-wide.xlsx xlsxspread:1 20 200
taskset -c 18 perf record -e cycles:u -F 4999 -o p.data -- \
  probe0624 savebench dense-wide.xlsx xlsxspread:1 40

# the scaling curve on a dumped changed set
taskset -c 18 probe0624 dump dense-wide.xlsx xlsxspread:1 set/
taskset -c 18-25 probe0624-deflate files set/ 10 40 1 2 4 8
```

The raw `perf.data` files are **not** retained: they are 43 MB for one scenario
and nothing in the record cites anything but the symbol report, which is
retained in full above the 0.15% cut with the summed share appended.

## What is not here

* No before/after legs. This change implements nothing, so there is no after.
* No end-to-end parallel save timing. None exists to take; the record's
  end-to-end figures are Amdahl arithmetic and are labelled modelled.
* No allocation or peak-RSS measurement. Gate A5 of the record makes those an
  admission requirement for an implementation; they were not taken here.
