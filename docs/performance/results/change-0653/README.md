# change 0653 — retained evidence

Record:
[`../../0653-mce-namespace-emission-rewrite.md`](../../0653-mce-namespace-emission-rewrite.md).
`performance_claim: none`. The markup-compatibility writer stops re-declaring
every in-scope namespace on every element it emits; the slicing consumers
re-declare at the slice boundary instead. Authorized by decision 1 of change
[0652](../../0652-owner-decisions-for-the-third-wave.md).

## Provenance

| | |
|---|---|
| Base commit | `70d7768cc6dada420ede063f72c88dc99ad30383` (branch `feat/office-format-completeness`) |
| Branch | `perf/0653-mce-namespace-emission-rewrite` |
| Before checkout | `/home/zhuhe/code/litchi-worktrees/before-70d7768cc` (untouched detached checkout of the base) |
| After worktree | `/home/zhuhe/code/litchi-worktrees/0653` |
| Host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| Toolchain | rustc 1.95.0 (59807616e 2026-04-14), valgrind 3.26.0, perf |
| Build | `--release --locked`, `debug = 1` for the probes; every measured process pinned with `taskset -c 8` |
| Concurrency | seven other measurement agents were active for most of this window; the load average was 25 to 45 on 32 cores while the deterministic counts, the deck, the XLSX timing and change 0638's harness rows were taken, and 9.5 while the `timing/window3/` and `marker/` legs were. Every quoted A/A floor is the one earned in its own window |

Every binary was staged outside its Cargo target directory before it ran
(change 0627's lesson) in `/home/zhuhe/code/litchi-worktrees/0653-bin/`:

| binary | sha256 |
|---|---|
| `xmlprobe.before` | `a5c9d9fdc28f36f06158f90f1e81a9ee5d1d444d1bb0a95c8c9b7e7933cfbbb7` |
| `xmlprobe.after` | `9424dd7bba68669a81a606ce931319f4c2517d82a0ef6601df989fabdf0f35f3` |
| `probe0649.before` | `53f156852599ee4dbefc09299f53adf68ff71c025b151479594307b13d30eae3` |
| `probe0649.after` | `aa244c669ae3efcbfe2d9ac1ea92b7925a42fe4cc8da3e56bb471f7b935a9532` |
| `litchi-perf-baseline.before` | `789e54ea451722bf1a7f90a6ca8ccf6c36c94eac4d34c4172da591b50456a833` |
| `litchi-perf-baseline.after` | `40504322661e12a68d92ff67e048b96cbaa212165ec608d85d6989b16719996f` |
| `marker-baseline.before` (base crates + change 0664's `tools/` diff) | `4fb8aa3ddb635e3937ec259e9603e80d4ea92ccef956a249ceae5ef8f993a0dd` |
| `marker-baseline.after` (this change + change 0664's `tools/` diff) | `33cf012ffdb7f4dcb2e7e2d5a5c03f205b84c4bbe0ae8842e7e33bb37dc234e3` |
| `sinkprobe.before` | `b2ead2ef5a554fc50acdee3935ffc3d1dfec90497ac5017e3ba3fc31526d60ef` |
| `sinkprobe.after` | `7424e070f2207b18d7249c4c77781c2329ddf05ed1304281e14ac01620df4eeb` |
| `readprobe.before` | `e6293faf7705f157a23c029f0eb5ed3c1fabb1bf29300ac060dc8e9073eb39f3` |
| `readprobe.after` | `938665736c8e986fd676d7d175f33218dcfc5c92329b13abe31f7a5a1236d441` |
| `saveprobe`, before | `c7a8c32f14cdc4c8453aaed2bcc0a98b108ba3ff833143753bdffc1f5ba35254` |
| `saveprobe`, after | `098c625f64193261794676ba0dca6f87ab097a6671e98a1f3571a809da3f76e3` |

Fixtures: `test-data/poi/test-data/spreadsheet/Excel_file_with_trash_item.xlsx`
("real", `sheet1.xml` 209,931 B), change 0587's
`results/change-0587/xml-substrate/control.xlsx` ("control", the same package
with only the worksheet's `mc`/`x14ac` markers stripped), and change 0649's deck
`test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx` with
the marker-stripped control `deck/run-phases.sh` and change 0649's
`scripts/run.sh` build in six lines of Python. Cell read: `H680`. Both legs read
the **before** checkout's `test-data`, so the inputs are byte-identical.

## Contents

| path | what it is |
|---|---|
| `probe/xmlprobe-main.rs` | the scratch probe (change 0588's, unchanged). Modes: `eager`/`source` (public XLSX read of one cell), `mce` (codec `n` times, prints in/out sizes), `raw` (exact output length, FNV-64 hash, borrowed flag, `Report`, or the refusal's `Debug` + `Display`), `canon` (namespace-resolving projection), `time` (one wall-clock nanosecond sample per line). Its manifest is change 0588's `results/change-0588/probe/Cargo.toml` with the path dependencies pointed at each leg |
| `oracle/mce_oracle_0653.py` | the corpus differential: every `.xlsx`/`.docx`/`.pptx` under `test-data`, every `.xml`/`.rels` member, before against after, comparing the resolving projection, the `Report`, the refusal identity, the borrow decision and the output length |
| `oracle/oracle-0653.tsv` | its report: `fixtures=320 parts=6964 refusals=0 mismatches=0 grew=0 shrank=980 before_output_bytes=197811418 after_output_bytes=36761542` |
| `oracle/mce_mutate.py`, `oracle/make-seeds.py`, `oracle/seeds/*.xml` | change 0588's adversarial differential, its seed generator and the five synthetic seeds (the four large seeds are real fixture parts, regenerated rather than copied) |
| `oracle/mutation-canon.tsv` | its report: `seeds=9 mutants=30000 refusals=16661 mismatches=0` |
| `counts/capture-cg-0653.sh` | the callgrind capture, one leg per invocation, `mce` (isolation pairs) or `public` (whole-process eager and source-backed reads) |
| `counts/cg-before/`, `counts/cg-after/` | per run: `*.inclusive.txt` and `*.self.txt` (annotated symbol tables; the raw `.out` files are deleted after extraction) and `*.stdout` (the probe's own in/out byte report). `mce-*-r1`/`mce-*-r6` are the isolation pair; the rest are whole-process single runs |
| `timing/time-probe-0653.sh`, `timing/rerun-eager-real.sh`, `timing/run-all.sh` | the paired-timing drivers (ABBA plus dedicated A/A blocks in the same window) and the sequencing script for the whole capture |
| `timing/stats-0653.py` | the percentile summary; its A/A floor is the widest p50 spread over **every** before-leg block, not one favourable pair |
| `timing/samples/*.txt` | every raw nanosecond sample, one file per block |
| `timing/summary.txt` | p50/mean/p95/p99 per leg, paired deltas in both directions, and the floor |
| `deck/run-phases.sh`, `deck/summarize.py` | change 0649's real-deck edit, both legs, four blocks per leg plus two A/A blocks |
| `deck/{real,control}.{before,after,floor}.N.tsv` | the per-block p50/mean/p95/min of each of the five documented calls and the edit total |
| `deck/summary.txt` | the per-phase table the record quotes |
| `harness/run.sh`, `harness/summarize.py` | the `litchi-perf-baseline` ABBA: change 0588's four marker-free XLSX selectors (the no-regression leg) and change 0638's three real-file PPTX ordinary-save rows on the 0649 deck |
| `harness/{x,p}-*.json` | the twelve harness reports (before x2, after x2, floor x2, for each family) |
| `harness/summary.txt` | per-selector, per-shape p50 before/after with the A/A floor |
| `harness/isolated/` | `pptx_real_file_ordinary_save_counting_publish` alone, its own process, 50 samples, six blocks per leg, plus `perfstat.txt`, the cycles/instructions isolation pair (samples 20 against 120) that shows the per-sample work fell 83% while the timed publish window rose |
| `publication/saveprobe/` | the corpus publication probe: twelve documented save routes over all 320 fixtures, one row per published archive member |
| `publication/rows-{before,after}.tsv.gz` | its raw rows, 20,421 members per leg |
| `publication/compare.py`, `publication/summary.txt`, `publication/refusals.tsv` | the comparison, every number the record quotes recomputed from the rows, and the outcome rows side by side |
| `marker/run.sh`, `marker/summarize.py` | change 0664's marker-bearing DOCX and PPTX selectors and their byte-identical marker-stripped controls, ABBA plus two A/A blocks, on both legs. 0664's harness diff (commit `e76897b0e`, `tools/` only) was applied to a fresh detached checkout of the base and to this worktree to build the two binaries, and reverted afterwards |
| `marker/*.json` | the eighteen harness reports |
| `marker/summary.txt` | per-selector p50 with the floor, and the marker/control ratio on each leg |
| `sink/sinkprobe/`, `sink/make-sink-fixtures.py` | the witness for change 0664's DOCX text-sink refusal: one `word/document.xml` with a 33-declaration root and 200 paragraphs, and the same bytes with the markup-compatibility URI replaced by an inert URI of the same length |
| `sink/result.txt` | its outcome on both legs, and the two probe SHA-256s |
| `timing/window3/`, `timing/rerun2-eager-real.sh` | the eager-real scenario re-measured in a quiet window after the first two were destroyed by unpinned neighbours; both failed windows stay under `timing/samples/` |
| `reads/readprobe/` | the public-read probe: every documented read of the three format crates over the same corpus |
| `reads/digest-{before,after}.txt.gz` | its raw digests, 179,205 observation lines per leg |
| `reads/dump-{before,after}.txt.gz` | the base64 markup each accessor returned, needed to recompute the canonical-equality classes |
| `reads/diff.tsv`, `reads/compare.py`, `reads/summary.txt`, `reads/binaries.sha256` | every differing line with its class, the script that classified them (including four negative controls), the numbers, and the binary hashes |
| `gates.txt` | the tail of every gate |
| `decision.json` | the decision record |
| `log-sections.md` | the four program-log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md` (the coordinator merges those files) |
| `cleanup.json` | what was deleted and what was kept |

## Replay

```sh
# one worktree per leg, each with its own CARGO_TARGET_DIR
git -C <repo> worktree add -b perf/0653-... <after> 70d7768cc   # then apply the change

# probes: change 0588's probe/Cargo.toml and change 0649's probe/Cargo.toml.example,
# with <checkout> replaced by each leg, built with CARGO_TARGET_DIR outside the repo
# and the binary copied out of it before it is run.

# deterministic counts
counts/capture-cg-0653.sh <leg> <bin>/xmlprobe.<leg> <scratch> 8 mce
counts/capture-cg-0653.sh <leg> <bin>/xmlprobe.<leg> <scratch> 8 public

# differentials
python3 oracle/make-seeds.py <repo>/test-data <scratch>/seeds
cp oracle/seeds/*.xml <scratch>/seeds/
python3 oracle/mce_oracle_0653.py <bin>/xmlprobe.before <bin>/xmlprobe.after \
        <repo>/test-data oracle-0653.tsv
python3 oracle/mce_mutate.py <bin>/xmlprobe.before <bin>/xmlprobe.after \
        <scratch>/seeds 30000 mutation-canon.tsv canon
python3 publication/compare.py   # see its header for the probe legs it expects
python3 reads/compare.py

# timing, last and pinned
timing/run-all.sh <scratch> <bin> 8
harness/run.sh <scratch>/harness <bin>/litchi-perf-baseline.before \
        <bin>/litchi-perf-baseline.after <deck> 8
```

Both Python differentials drive **thread** pools over subprocesses: this host's
Python 3.14 breaks `ProcessPoolExecutor` for a script with top-level code, and
every unit of work here is a blocking subprocess call, so threads lose nothing.
