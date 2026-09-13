# ODS selector performance evidence

This report compares an absolute baseline at commit `d5b5afdb4032b718c0f94c113a3b1b09f17ae6ba` with the candidate binary-search change in `crates/litchi-ods/src/sheet_metadata/index.rs` (candidate SHA256 `6f08a6488414c1eb357fa6c877ffe41d3d7c27eed9682ee4f62b3e815b10aa2b`). The candidate is reconstructed by applying [`candidate/index.diff`](candidate/index.diff) to that exact commit; the patch SHA256 is recorded in [`candidate/index.diff.sha256`](candidate/index.diff.sha256).

Both binaries use the same corrected retained harness, explicit ODF range endpoints, one synthetic sheet, three warmups, and fifteen measured iterations. Each profiled process was pinned to CPU 2. Duration and budget-counter rows are per measured iteration; the `probes` column is the cumulative successful operation count across the fifteen measured iterations, while max RSS is the process high-water value across the run. The host was shared with background services and a concurrent root gate/build workload; this adds run-to-run noise and means the numbers are host-constrained observations, not isolated machine capacity.

`parse` includes XML parsing in its timer. `lookup` parses before timing and visits every position in the selected grid. `stage-one` parses before timing and stages one cell. `stage-batch` parses before timing and stages every cell. `edit-batch` includes parsing, staging, commit, and the harness readback path. All selector operations use position selectors against dense synthetic physical cells; no claim is made here about worksheet-name lookup or merge fallback.

`Memory` is the harness `Resource::Memory` reservation counter in bytes and `Work` is its `Resource::Work` counter in units. These are budget-accounting values, not allocator peak measurements. For lookup/staging, the delta excludes parse while final includes the whole iteration context. `max RSS` is the operating system maximum resident set size from `/usr/bin/time -v`.

| lane | baseline p50 → candidate p50 (ms) | p50 time change | Work Δ baseline → candidate | Memory final p50 baseline → candidate | max RSS baseline → candidate (KiB) | probes |
|---|---:|---:|---:|---:|---:|---:|
| parse 64×64 | 5.968 → 5.956 | -0.20% | 1,197,402 → 1,197,402 | 7,909,580 → 7,909,580 | 9,276 → 9,288 | 0 → 0 |
| lookup 32×32 | 0.314 → 0.117 | -62.78% | 278,528 → 77,312 | 1,990,540 → 1,990,540 | 4,364 → 4,404 | 15,360 → 15,360 |
| lookup 64×64 | 2.313 → 0.531 | -77.05% | 2,162,688 → 368,640 | 7,909,580 → 7,909,580 | 9,400 → 9,388 | 61,440 → 61,440 |
| lookup 256×256 | 138.220 → 11.978 | -91.33% | 135,266,304 → 7,905,280 | 126,042,188 → 126,042,188 | 109,500 → 109,540 | 983,040 → 983,040 |
| stage-one 64×64 | 0.001 → 0.001 | +55.38% | 88 → 376 | 7,910,632 → 7,910,632 | 9,388 → 9,400 | 15 → 15 |
| stage-batch 32×32 | 1.224 → 0.595 | -51.42% | 851,968 → 248,320 | 3,067,788 → 3,067,788 | 4,616 → 4,636 | 15,360 → 15,360 |
| stage-batch 64×64 | 8.159 → 2.677 | -67.19% | 6,553,600 → 1,171,456 | 12,218,572 → 12,218,572 | 10,276 → 10,256 | 61,440 → 61,440 |
| edit-batch 32×32 | 9.338 → 8.009 | -14.23% | 3,003,519 → 1,796,223 | 7,962,392 → 7,962,392 | 10,760 → 10,764 | 15,360 → 15,360 |
| edit-batch 64×64 | 43.455 → 32.462 | -25.30% | 18,277,823 → 7,513,535 | 31,744,408 → 31,744,408 | 34,780 → 34,568 | 61,440 → 61,440 |

The position-lookup and batch-staging lanes show lower timed Work and p50 latency for this dense synthetic grid after the candidate change. The small control lanes expose its fixed search overhead: `stage-one` is 650 ns baseline versus 1,010 ns candidate (+55.38%) and Work Δ rises from 88 to 376 units. The 64×64 parse control is 5.968 ms versus 5.956 ms (−0.20%), with equal budget counters. RSS changes are small in these runs (−212 to +40 KiB); they do not establish an allocator or peak-memory improvement.

The 256×256 lookup is a bounded stress point, with 983,040 successful probes per run set. Its baseline p50 Work Δ is 135,266,304 units and candidate p50 Work Δ is 7,905,280 units; both runs completed under the harness profile. This is evidence for the selected workload only and should not be generalized to all ODS documents or all selector paths.

Both runs exited zero on all nine lanes. The complete per-lane stdout and `/usr/bin/time -v` records are retained in [`baseline/`](baseline/) and [`candidate/`](candidate/); [`comparison.csv`](comparison.csv) contains the machine-readable comparison. Baseline and candidate source manifests each pass their recorded SHA256 checks. The candidate build was performed with the same repository-relative harness manifest and a separate external target.

Replay the candidate with a temporary worktree rooted at the pinned commit and the retained patch (the captured build commands are retained verbatim in each command ledger):

```sh
git worktree add --detach /var/tmp/ods-selector-replay d5b5afdb4032b718c0f94c113a3b1b09f17ae6ba
git -C /var/tmp/ods-selector-replay apply --unidiff-zero /absolute/path/to/ods-selector-performance/candidate/index.diff
CARGO_TARGET_DIR=/var/tmp/ods-selector-candidate-target cargo build --manifest-path /var/tmp/ods-selector-replay/docs/report/spec-gap-validation-evidence/ods-sheet-metadata/harness/Cargo.toml --release --locked --offline
taskset -c 2 /usr/bin/time -v -o candidate.time /var/tmp/ods-selector-candidate-target/release/ods-meta-profile --workload lookup --sheets 1 --rows 64 --columns 64 --warmups 3 --iterations 15
git worktree remove --force /var/tmp/ods-selector-replay
```

The exact nine command lines are in [`candidate/commands.txt`](candidate/commands.txt); the baseline command ledger is [`baseline/commands.txt`](baseline/commands.txt).
