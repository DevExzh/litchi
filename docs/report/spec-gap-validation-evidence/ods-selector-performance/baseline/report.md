# ODS selector performance baseline

This is an absolute baseline run for the selector optimization, pinned to commit `d5b5afdb4032b718c0f94c113a3b1b09f17ae6ba` (`feat(ods): add bounded sheet metadata transactions`). The candidate comparison, if captured, must use the same corrected harness and lane arguments.

All lanes used one synthetic sheet, three warmups, and fifteen measured iterations. The XML fixture has explicit ODF range endpoints (`Sheet0.A1:Sheet0.A1`) and the harness source is the already retained corrected fixture at [`../../ods-sheet-metadata/harness/src/main.rs`](../../ods-sheet-metadata/harness/src/main.rs). The process was pinned to CPU 2 with `taskset`; the host was shared with background processes and a concurrent root gate/build workload, so these are host-constrained observations rather than isolated machine capacity.

`parse` includes XML parsing in its timer. `lookup` and `stage-*` parse before the timer: lookup visits every cell in the selected grid, `stage-one` stages one cell, and `stage-batch` stages every cell. `edit-batch` includes parse, staging, commit, and the harness readback path. `probes` is the total successful operation count accumulated across the fifteen measured iterations; duration and budget-counter fields are per-iteration samples, while max RSS is the process high-water value across the run.

`Memory` and `Work` are the harness-reported `Resource::Memory` reservation bytes and `Resource::Work` units. They are budget counters, not allocator peak measurements. `memory_delta_p50`/`work_delta_p50` describe the timed operation where the harness excludes parse for lookup/staging; `memory_final_p50`/`work_final_p50` include the whole per-iteration context. `max_rss_kib` is the operating system maximum resident set size from `/usr/bin/time -v`.

| lane | source bytes | p50 (ms) | p95 (ms) | Memory Δ p50 (budget B) | Memory final p50 (budget B) | Work Δ p50 (units) | Work final p50 (units) | max RSS (KiB) | probes |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| parse 64×64 | 338,634 | 5.968 | 6.265 | 7,909,580 | 7,909,580 | 1,197,402 | 1,197,402 | 9,276 | 0 |
| lookup 32×32 | 85,610 | 0.314 | 0.314 | 0 | 1,990,540 | 278,528 | 580,858 | 4,364 | 15,360 |
| lookup 64×64 | 338,634 | 2.313 | 2.319 | 0 | 7,909,580 | 2,162,688 | 3,360,090 | 9,400 | 61,440 |
| lookup 256×256 | 5,383,434 | 138.220 | 138.387 | 0 | 126,042,188 | 135,266,304 | 154,306,458 | 109,500 | 983,040 |
| stage-one 64×64 | 338,634 | 0.001 | 0.001 | 1,052 | 7,910,632 | 88 | 1,197,490 | 9,388 | 15 |
| stage-batch 32×32 | 85,610 | 1.224 | 1.253 | 1,077,248 | 3,067,788 | 851,968 | 1,154,298 | 4,616 | 15,360 |
| stage-batch 64×64 | 338,634 | 8.159 | 8.170 | 4,308,992 | 12,218,572 | 6,553,600 | 7,751,002 | 10,276 | 61,440 |
| edit-batch 32×32 | 85,610 | 9.338 | 9.471 | 7,962,392 | 7,962,392 | 3,003,519 | 3,003,519 | 10,760 | 15,360 |
| edit-batch 64×64 | 338,634 | 43.455 | 44.897 | 31,744,408 | 31,744,408 | 18,277,823 | 18,277,823 | 34,780 | 61,440 |

The baseline scaling observations are visible in the retained raw rows: timed lookup work is 278,528, 2,162,688, and 135,266,304 units for 32×32, 64×64, and 256×256; stage-batch work is 851,968 and 6,553,600 units for 32×32 and 64×64. These values describe this synthetic sparse-grid workload and do not support a broader performance claim.

Reproduction uses the repository-relative harness manifest after checking out the pinned commit:

```sh
git checkout d5b5afdb4032b718c0f94c113a3b1b09f17ae6ba
CARGO_TARGET_DIR=/var/tmp/ods-selector-baseline-target cargo build --manifest-path docs/report/spec-gap-validation-evidence/ods-sheet-metadata/harness/Cargo.toml --release
taskset -c 2 /usr/bin/time -v -o baseline.time /var/tmp/ods-selector-baseline-target/release/ods-meta-profile --workload lookup --sheets 1 --rows 64 --columns 64 --warmups 3 --iterations 15
```

The full lane command ledger is in [`commands.txt`](commands.txt), raw harness records in [`raw.csv`](raw.csv) and the per-lane `.out` files, and source provenance in [`commit.txt`](commit.txt) plus [`source-hashes.sha256`](source-hashes.sha256).
