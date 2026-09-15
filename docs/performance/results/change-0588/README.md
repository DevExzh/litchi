# change 0588 — retained evidence

Record: [`../../0588-mce-codec-namespace-emission.md`](../../0588-mce-codec-namespace-emission.md).
`performance_claim: none`. The implemented subset is **byte-identical**: the MCE
codec's processed output, its `Report` counters, its borrow-versus-own decision
and every refusal identity are unchanged. The namespace-emission rewrite that
survey item XML-1 asked for is designed, measured and **withdrawn**; its patch
and its own differential reports are under `design/`.

## Provenance

| | |
|---|---|
| Base commit | `08d968f8ec7db27cf1187d01911fd08b9d014d91` (branch `feat/office-format-completeness`) |
| Branch | `perf/0588-mce-codec-namespace-emission` |
| Before checkout | `/home/zhuhe/code/litchi-worktrees/before-08d968f8e` (untouched detached checkout of the base) |
| After worktree | `/home/zhuhe/code/litchi-worktrees/0588` |
| Host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| Toolchain | rustc 1.95.0 (59807616e 2026-04-14), valgrind 3.26.0 (callgrind) |
| Build | `--release --locked`, `debug = 1` for the probe; every measured process pinned with `taskset -c 8` |
| Concurrency | seven other measurement agents were active on the host; the A/A floor was measured in the same window |

Binary identity:

| binary | bytes | sha256 |
|---|---:|---|
| `xmlprobe` (before leg) | 83,272,360 | `e64a54c0ea12f2a0a179baa6844c55f772d88f6d695f97dc15591b397104e0a3` |
| `xmlprobe` (after leg) | 83,239,440 | `374934f85e03359dc10c106803e0ad608c6d57b041ff1e240f8e0c7fec08c7d0` |
| `litchi-perf-baseline` (before leg) | 60,112,744 | `0bf9d22f468a58c1a2754b8c8ef2c46ec00b7879fb87038d4c4b38dd6b409c9b` |
| `litchi-perf-baseline` (after leg) | 60,097,272 | `8d779a084f3458ae82d856160b64a196764cb14f0da49eb186d34378151e05ec` |

Fixtures: `test-data/poi/test-data/spreadsheet/Excel_file_with_trash_item.xlsx`
("real", `sheet1.xml` 209,931 B, 11,578 elements, 681 `x14ac:dyDescent`) and
change 0587's `results/change-0587/xml-substrate/control.xlsx` ("control", the
same package with only the `mc`/`x14ac` worksheet markers stripped). The two
packages still both declare the MCE namespace in `xl/workbook.xml` and
`xl/styles.xml`, which is why the control also improves. Cell read: `H680`.

## Contents

| path | what it is |
|---|---|
| `probe/Cargo.toml`, `probe/main.rs` | the scratch probe, extended from change 0587's `xmlprobe`. Modes: `eager`/`source` (public XLSX read of one cell), `mce` (codec `n` times, prints in/out sizes), `raw` (byte-identity oracle: exact output length, FNV-64 content hash, borrowed flag, `Report`, or the refusal's `Debug` and `Display`), `canon` (namespace-resolving projection), `time` (one wall-clock nanosecond sample per line). Path dependencies point at the leg's checkout. |
| `oracle/mce_oracle.py` | corpus differential: every `.xlsx`/`.docx`/`.pptx` under `test-data`, every `.xml`/`.rels` member, before against after, in `raw` or `canon` mode |
| `oracle/mce_mutate.py` | adversarial differential: single-byte substitutions, deletions, insertions and truncations over MCE-bearing seeds, deterministic per index |
| `oracle/make-seeds.py` | regenerates the four large seeds (real fixture parts) that are not copied here |
| `oracle/seeds/*.xml` | the five synthetic seeds, retained verbatim |
| `oracle/oracle-raw.tsv` | corpus, byte identity: `fixtures=320 parts=6964 mismatches=0` |
| `oracle/oracle-canon.tsv` | corpus, namespace-resolving identity: `fixtures=320 parts=6964 mismatches=0` |
| `oracle/mutation-raw.tsv` | mutants, byte identity: `mutants=30000 refusals=16661 mismatches=0` |
| `counts/capture-cg.sh` | the callgrind capture, one leg per invocation |
| `counts/cg-before/`, `counts/cg-after/` | per run: `*.inclusive.txt` and `*.self.txt` (annotated symbol tables, raw `.out` files deleted after extraction) and `*.stdout` (the probe's own in/out byte report). `mce-*-r1`/`mce-*-r6` are the isolation pair; `eager-*`/`source-*` are whole-process single runs |
| `counts/summary.txt` | the differenced per-call and whole-process figures |
| `timing/time-probe.sh`, `timing/stats.py` | the paired-timing driver (ABBA plus an A/A block in the same window) and its percentile summary |
| `timing/samples/*.txt` | all 960 raw nanosecond samples, one file per block |
| `timing/summary.txt` | p50/mean/p95/p99 per leg, paired deltas in both directions, and the A/A floor |
| `harness/run.sh`, `harness/summarize.py` | the `litchi-perf-baseline` ABBA over four existing XLSX selectors (the no-regression leg; the harness corpora carry no MCE markers) |
| `harness/*.json` | the six harness reports (before x2, after x2, floor x2) |
| `harness/summary.txt` | per-selector p50 before/after with the A/A floor |
| `gates.txt` | the tail of every gate, plus the pre-existing `litchi-iwa` example failure reproduced on the untouched base checkout |
| `decision.json` | the decision record |
| `log-sections.md` | the four log paragraphs for HOTSPOTS, GOAL_AUDIT, REPORT and ADR_COMPLIANCE (the coordinator merges those files) |
| `design/namespace-emission.patch` | the **withdrawn** namespace-emission rewrite, complete with its nine unit tests |
| `design/oracle-canon-withdrawn.tsv` | that patch against the corpus, namespace-resolving: `parts=6964 mismatches=0` |
| `design/mutation-canon-withdrawn.tsv` | that patch against 30,000 mutants, namespace-resolving: `refusals=16661 mismatches=0` |
| `design/docxprobe/` | a scratch probe built against the **untouched base checkout**, with `RESULT.txt`: `Paragraph::extensions()` already refuses on the one `test-data` `.docx` whose `word/document.xml` carries no MCE namespace, with the identical error the withdrawn rewrite produces everywhere |

## Replay

```sh
# one worktree per leg, each with its own CARGO_TARGET_DIR
git -C <repo> worktree add -b perf/0588-... <after> 08d968f8e   # then apply the change
# probe (adjust the path dependencies in probe/Cargo.toml to the leg's checkout)
CARGO_TARGET_DIR=<targets>/0588-<leg> cargo build --release

# deterministic counts
counts/capture-cg.sh <leg> <targets>/0588-<leg>/release/xmlprobe <scratch> 8

# differentials
python3 oracle/make-seeds.py <repo>/test-data <scratch>/seeds
python3 oracle/mce_oracle.py <before-bin> <after-bin> <repo>/test-data out.tsv raw
python3 oracle/mce_oracle.py <before-bin> <after-bin> <repo>/test-data out.tsv canon
python3 oracle/mce_mutate.py <before-bin> <after-bin> <scratch>/seeds 30000 out.tsv raw

# paired timing (real fixture) and the no-regression harness leg
timing/time-probe.sh <scratch> <before-bin> <after-bin> 8 40 && python3 timing/stats.py <scratch>/timing
harness/run.sh <scratch> <before-harness> <after-harness> \
  xlsx_open_owned,xlsx_first_cell,xlsx_full_cell_scan,xlsx_narrow_column_range_scan 8
python3 harness/summarize.py <scratch>/harness
```

## What is not here

No allocator report, no `perf stat` cycle counts, no RSS or cold-cache
measurement, no DOCX or PPTX public-read profile, and no corpus-wide
distribution of the improvement. See the record's Limitations section.
