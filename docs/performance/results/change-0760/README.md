# Evidence for change 0760

Record: [`../../0760-pptx-slide-root-memo-reapply.md`](../../0760-pptx-slide-root-memo-reapply.md).
Base `1d1044e3ac`; measured head `dfde1e43bb` on
`perf/0760-pptx-slide-root-memo-reapply` (the slide-root memo of change 0743's
withdrawn `99ce9c5e34`, re-applied under ADR 0032). `performance_claim: none`.

## Contents

| path | what |
| --- | --- |
| `timing/` | the before/after harness run: 48 raw harness JSON reports (`r<round>-<group>-p<position>-<arm>.json`), a `perf stat` CSV per process (user instructions and cycles of the whole process), `run-log-*.json` (command, exit code, load average before each process, wall time), `analysis.json` and `tables.md` |
| `aa/` | the A/A floor: the before binary against a byte-identical copy at a path of equal length, two ABBA blocks, same layout |
| `counters/` | per-region user instructions and cycles from the probes: `commit-capture-control/` (`Transaction::commit`, capture, full text) and `cycles/` (whole edit/save cycles and their setup), each with `probe-counters.json` and every `perf stat` CSV; `counter-table.md` is the rendered table |
| `alloc/probe-alloc.txt` | exact allocation calls and requested bytes per region, base and head, two repeats each |
| `outputs/identity.md` | digests of the 30 artifacts (published archives, durable patches, revisions, full text; tiny, medium, large; no-op, one edit, one percent) written by base and head; all identical |
| `differential/README.md` | the base-versus-head differential over the 78 PPTX fixtures and 67,712 mutated packages: method, output digests (identical) and outcome counts |
| `profiles/commit-branch-fp.txt` | text summary of a frame-pointer profile of the remaining one-edit commit at the head |
| `probe/` | the probe source (`LITCHI_ROOT` stands for the checkout it is built against) |
| `measure.py`, `analyze.py`, `probe_counters.py`, `counter_table.py`, `gate.sh` | the drivers used |
| `binaries.sha256` | every measured binary |
| `gates.txt` | every gate command at the measured head with its exit code, test totals and output tail, plus the base comparisons |
| `cleanup.json` | what was removed and what was kept |
| `log-sections.md` | ready-to-paste sections for `HOTSPOTS.md`, `REPORT.md` and `GOAL_AUDIT.md` |

## Reproduction

The harness is unchanged. Both legs are built with the same command from
detached sparse worktrees of equal path length (the sparse checkout omits only
`docs/report/spec-gap-validation-evidence/`, which no build reads),
`0760-before-src` at `1d1044e3ac` and `0760-branch-src` at `dfde1e43bb`:

```text
CARGO_TARGET_DIR=<targets>/0760-before cargo build --release --locked --offline \
  --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline   # in 0760-before-src
CARGO_TARGET_DIR=<targets>/0760-branch cargo build ...                        # same, in 0760-branch-src
```

The binaries run from `bin/a/`, `bin/b/` and (the A/A copy) `bin/c/`, so every
argv has the same length. Then, each invocation in the foreground:

```text
python3 measure.py --before bin/a/litchi-perf-baseline --after bin/b/litchi-perf-baseline \
  --out timing --rounds 2 --round-offset 0 --core 4
python3 measure.py ... --out timing --rounds 2 --round-offset 2 --core 4
python3 measure.py --before bin/a/litchi-perf-baseline --after bin/c/litchi-perf-baseline \
  --out aa --rounds 2 --core 4
python3 analyze.py --runs timing --out timing
python3 analyze.py --runs aa --out aa
python3 probe_counters.py --a bin/pa/pptx-memo-probe --b bin/pb/pptx-memo-probe --out counters/... --core 4 --modes ...
python3 counter_table.py counters/*/probe-counters.json
```

Groups: `large` and `medium` run `pptx_semantic_full_text` (control),
`pptx_semantic_noop_edit_save`, `pptx_semantic_one_edit_save` and
`pptx_semantic_one_percent_edit_save` (15 samples after 3 warmups for large, 60
after 10 for medium); `phases` runs `pptx_semantic_opened_transaction_phases` on
both decks (15 after 3). Each block is before, after, after, before. Pairs are
(position 0, position 1) and (position 3, position 2) of each block; the
bootstrap resamples the paired processes 10,000 times with seed 760.

The probe builds the harness's own deterministic corpus (the generator is
copied from `tools/perf-baseline`) and calls exactly the public methods the
harness times. Build it with `cargo build --release --offline` after replacing
`LITCHI_ROOT` and copying `tools/perf-baseline/Cargo.lock` beside its manifest;
`--features count-alloc` for allocation counts. Modes: `commit <shape> <iters>
one|pct` (only `Transaction::commit`, the staged transaction cloned outside the
clock), `capture <shape> <iters>`, `setup <shape> <iters>`, `cycle <shape>
<iters> noop|one|pct`, `fulltext <shape> <iters>`, `alloc <shape> 1`,
`dump <shape> 1 <directory>` and `diff <fixture-root> <random-per-fixture>
<seed>`. `probe_counters.py` differences two iteration counts per mode and
`counter_table.py` subtracts `setup` from the `cycle` modes.

## Host and window

AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws; builds with the
worktree's pinned toolchain 1.95.0. Every measured process pinned to core 4.
Other agents built and measured on the host throughout (one-minute load average
11–29 before the A/B processes). No measurement ran while this change's own
builds ran. One A/B process (round 3, large group, position 2) was time-shared:
single samples of 129 ms and 65 ms against p50s of 28 and 24 ms, while its user
cycles matched every other after process; it is kept, and it is the source of
five of the seven regression flags, all on p95 or mean. No process was
discarded.
