# Evidence for change 0743

Record: [`../../0743-pptx-semantic-text-and-edit-path.md`](../../0743-pptx-semantic-text-and-edit-path.md).
Base `009d515bef`; measured head `ad2c490ee3` on
`perf/0743-pptx-semantic-text-and-edit-path` (the retained 0743 changes, the
revert of the slide-root memo, a compile-time guard, and change 0755's fix).
`performance_claim: none`.

The packet holds two series. **v2** measures the retained code and is the
record's primary evidence. **v1** measured the first head (`e74faad918`), which
still contained the slide-root memo now withdrawn and proposed as ADR 0032; it
is kept because the memo's measured effect is quoted from it.

## Contents

| path | what |
| --- | --- |
| `v2/timing/` | the primary before/after harness run: 48 raw harness JSON reports (`r<round>-<group>-p<position>-<arm>.json`), a `perf stat` CSV per process (user instructions and cycles), `run-log-*.json` (command, exit code, load average before each process, wall time), `analysis.json` and `tables.md` |
| `v2/aa/` | its A/A floor: the before binary against a byte-identical copy at a path of equal length, two ABBA blocks |
| `v2/aa-disturbed/` | the first A/A attempt, during which core 8 was time-shared; kept, not used as the floor |
| `v2/counters/` | per-operation user instructions and cycles from the probes (`probe-counters.json`, `region-counters.json` and every `perf stat` CSV) |
| `v2/alloc/probe-alloc.txt` | exact allocation calls and requested bytes per region, base checkout and head, two repeats each |
| `v2/outputs/identity.txt` | the 21 artifacts (published archives, revisions with patch emptiness, full text; tiny, medium, large; no-op, one edit, one percent) written by base and head, with digests; all identical |
| `timing-samecmd/`, `aa-samecmd/` | v1: the first head against a before leg built from the base checkout with the same command |
| `timing/`, `aa/` | v1: the first head against the prebuilt base harness |
| `attribution/` | v1: the probe built at each of the first series' ten commits (including the withdrawn memo), rotating order, three rounds |
| `alloc/`, `outputs/`, `open-counters/` | v1: allocation counts, output identity and open-path counters for the first head |
| `profiles/` | text summaries of the probe profiles that motivated the change (`fulltext`, `capture-fp`, `settext-fp`, `commit-pct-fp`, base code) and of the first head (`fulltext-fp3`, `capture-fp5`, `commit-pct-fp4`, `cycle-pct-fp3`); `perf record -F 2000–4000 -g` on a frame-pointer probe build, `--children` unless named `fulltext` |
| `probe/` | the probe source (`LITCHI_ROOT` stands for the checkout it is built against) |
| `measure.py`, `analyze.py`, `attribute.py`, `probe_counters.py`, `gate.sh` | the drivers used |
| `binaries.sha256` | every measured binary of both series |
| `gates.txt` | every gate command at the measured head with its exit code, test totals and output tail |
| `cleanup.json` | what was removed and what was kept |
| `log-sections.md` | ready-to-paste sections for `HOTSPOTS.md`, `REPORT.md` and `GOAL_AUDIT.md` |

## Reproduction (v2)

The harness is unchanged. Both legs are built with the same command from
detached worktrees of equal path length, `0743-src-a` at `009d515bef` and
`0743-src-b` at `ad2c490ee3`, each with the root `Cargo.lock` copied in:

```text
CARGO_TARGET_DIR=<targets>/0743/a cargo build --release --locked --offline \
  --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline   # in 0743-src-a
CARGO_TARGET_DIR=<targets>/0743/b cargo build ...                             # same, in 0743-src-b
```

The binaries run from `bin/a/`, `bin/b/` and (the A/A copy) `bin/c/`, so every
argv has the same length. Then, each invocation in the foreground:

```text
python3 measure.py --before bin/a/litchi-perf-baseline --after bin/b/litchi-perf-baseline \
  --out timing --rounds 2 --round-offset 0 --core 8 --counters
python3 measure.py ... --out timing --rounds 2 --round-offset 2 --core 8 --counters
python3 measure.py --before bin/a/litchi-perf-baseline --after bin/c/litchi-perf-baseline \
  --out aa --rounds 2 --core 8 --groups large,medium --counters
python3 analyze.py --runs timing --out timing
python3 probe_counters.py --a bin/pa/pptx-semantic-probe --b bin/pb/pptx-semantic-probe --out counters --core 8
```

Groups: `large` and `medium` run `pptx_semantic_open`, `pptx_semantic_full_text`,
`pptx_semantic_noop_edit_save`, `pptx_semantic_one_edit_save`,
`pptx_semantic_one_percent_edit_save` and the control `docx_semantic_full_text`
(15 samples after 3 warmups for large, 60 after 10 for medium); `phases` runs
`pptx_semantic_opened_transaction_phases` on both decks (15 after 3). Each
block is before, after, after, before. Pairs are (position 0, position 1) and
(position 3, position 2) of each block; the bootstrap resamples the paired
processes 10,000 times with seed 743.

The probe builds the harness's own deterministic corpus (same generator code)
and calls exactly the public methods the harness times: `presentation().text()`,
`Package::from_bytes`, and `opened_presentation_transaction` →
`set_shape_text` → `commit` → `apply_opened_presentation_commit` → `to_bytes`,
with verification outside the loop. Build it with `cargo build --release
--offline` after replacing `LITCHI_ROOT` and copying
`tools/perf-baseline/Cargo.lock` beside its manifest, and with `--features
count-alloc` for allocation counts. Modes: `fulltext <shape> <iters>`,
`open <shape> <iters>`, `setup <shape> <iters>`,
`cycle <shape> <iters> noop|one|pct`, `alloc <shape> 1`,
`dump <shape> 1 <directory>`. `probe_counters.py` differences two iteration
counts per mode and subtracts `setup` from the `cycle` modes.

## Host and window

AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws; builds with the
worktree's pinned toolchain 1.95.0 (the harness reports the default `rustc` on
the runtime `PATH`, which differs). Every measured process pinned to core 8.
Other agents built and measured on the host throughout. Core 8 was time-shared
during the first A/A attempt and in three A/B processes: those processes ran at
about twice the wall time of their pairs while their user cycles matched every
other process (see the record). No measurement ran while this change's own
builds ran. No process was discarded; the disturbed A/A attempt is kept beside
its rerun.

## Results

See the record for the tables. Each `tables.md` lists every paired ratio above
1.05 in p50, p95 or mean: seven in the v2 A/B run and three in its A/A floor,
all on p95 or mean tails of time-shared pairs, the 2 ms `pptx_semantic_open`
region, or the `docx_semantic_full_text` control — never on a p50 of a case
this change speeds up.
