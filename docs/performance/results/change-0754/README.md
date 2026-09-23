# Retained evidence — change 0754

Record: [`../../0754-docx-semantic-edit-and-text-path.md`](../../0754-docx-semantic-edit-and-text-path.md).

## Provenance

| | |
| --- | --- |
| base commit | `63ec6a5027` (branch tip with records 0742, 0744–0747 and 0750) |
| branch | `perf/0754-docx-semantic-edit-and-text-path`; final code `2c467b3e66` |
| before leg | detached worktree `/home/zhuhe/code/litchi-worktrees/0754-before-src` at the base, root `Cargo.lock` copied in, target `targets/0754-before`; removed after use |
| after leg | the branch worktree, target `targets/0754` (harness, tests, gates) |
| host | AMD EPYC 9R45, 32 cores, `Linux 7.0.0-1012-aws x86_64`, shared with other agents |
| CPU pin | `taskset -c 12` for every timed, perf-stat and callgrind process |
| harness build | `cargo build --release --manifest-path tools/perf-baseline/Cargo.toml --locked --offline --bin litchi-perf-baseline`, identical for both legs (`scripts/build_before.sh`); rustc 1.95.0 pinned by `rust-toolchain.toml`; binaries copied to equal-length paths `bin/lpb-A`, `bin/lpb-B` |
| probe | `probe/` built against each leg's checkout (`Cargo.toml.template` with `<CHECKOUT>` replaced, the harness's `Cargo.lock`), `cargo +1.95.0 build --release --offline`, debug line tables, a counting global allocator |

Binary SHA-256s are in [`binaries.txt`](binaries.txt).

## Contents

| path | what it is |
| --- | --- |
| `timing/final/` | the reported ABBA campaign (6 rounds, 12 cases, 288 processes) on `2c467b3e66`: `table.txt`, `analysis.json` (per-process p50/p95/mean, whole-process instructions and cycles, paired ratios, bootstrap CIs, digests), `flags.json` (every paired comparison moving more than 5%), `raw.tar.gz` (every JSON report, perf-stat file and log, and `status.txt`) |
| `timing/rerun-sb/` | `docx_source_backed_one_edit_save` alone, 8 rounds, final binaries |
| `timing/campaign-1/` | the intermediate campaign on `3ab983e5cf` (before the bounded index and the identity restriction), kept unedited |
| `timing/rerun-ordinary/` | `docx_ordinary_save_lifecycle` alone, 6 rounds, on the intermediate campaign's binaries |
| [`instructions/loops-table.md`](instructions/loops-table.md), `loops-summary.json` | per-iteration instructions, cycles, allocations and requested bytes of the timed operations (probe `loop`, 60 minus 10 iterations) |
| [`instructions/audits-per-iteration.txt`](instructions/audits-per-iteration.txt) | audit entry points' calls and inclusive instructions per one-edit iteration (callgrind, 3 minus 1 iterations) |
| [`instructions/worst-case-witnesses.txt`](instructions/worst-case-witnesses.txt) | `Document::text` on the five namespace worst-case witnesses, both legs |
| [`differential/`](differential/README.md) | the base-versus-branch differential: inputs, per-input outputs of both legs, comparison |
| `probe/` | the probe's source |
| `scripts/` | every script used: builds, ABBA runners, analysis, flags, differential inputs and comparison, witnesses, instruction loops, callgrind attribution, gate runner |
| [`gates.txt`](gates.txt) | every gate's command, exit status, duration and test totals on the final commit, and the mutation checks of the new tests |
| [`log-sections.md`](log-sections.md) | ready-to-paste sections for `HOTSPOTS.md`, `REPORT.md` and `GOAL_AUDIT.md` |
| [`cleanup.json`](cleanup.json) | what was removed and what was kept |

Not kept: binaries (hashes kept), generated corpora and differential inputs
(rebuilt by the harness generators, `scripts/make_inputs.py` and
`scripts/make_dos.py`), callgrind and perf data (summaries kept), build logs.

## Replay

```sh
# 1. before-leg source: git worktree add --detach <dir> 63ec6a5027; copy the root Cargo.lock
# 2. harness legs: scripts/build_before.sh, then the identical command on the branch; copy to bin/lpb-A, bin/lpb-B
# 3. timing: scripts/run_abba.sh; scripts/analyze.py <out>; scripts/flags.py <out>/analysis.json <out>/flags.json
#    (scripts/run_abba_sb.sh and run_abba_ordinary.sh for the single-case reruns, ROUNDS=8 / 6)
# 4. probe: probe/Cargo.toml.template per leg; `probe0754 gen <corpus dir>` writes the semantic corpora
# 5. differential: scripts/make_inputs.py <repo> <corpus dir> inputs; probe `docx`, `snapxml`, `pptx` modes
#    read file lists on stdin; scripts/compare.py in the output directory
# 6. witnesses: scripts/make_dos.py <semantic-tiny.docx> <dir>; `probe0754 loop text <witness> 3`
# 7. instructions: scripts/loops.sh <before probe> <after probe> <dir>; scripts/summarize.py <dir>;
#    callgrind of `probe0754 loop one <large corpus> {1,3}` per leg; scripts/cg_audits.py <dir>
```
