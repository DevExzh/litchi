# Evidence: change 0658, the selected-cell ineligibility gate

Change record:
[`0658-xlsx-selected-cell-ineligibility-gate.md`](../../0658-xlsx-selected-cell-ineligibility-gate.md).

Disposition: **retained, implemented**. `performance_claim: none`. The gate
change [0597](../../0597-xlsx-selected-cell-ineligibility-gate.md) froze is
landed under decision 8 of
[0652](../../0652-owner-decisions-for-the-third-wave.md), by a different
mechanism: the selected-cell scan stops the MCE stream at the event that settles
an ineligible verdict instead of running to EOF. An ineligible source-backed
one-cell read falls by 64.43% and 23.14% in instructions on the two fixtures
0597 priced; the whole-corpus transcript is byte-identical on both legs.

## Contents

| Path | What it is |
| --- | --- |
| `callgrind/` | Isolation-pair annotations (inclusive at `--threshold=99.99`, and self) for N = 1 and N = 11 source-backed one-cell reads, both legs, on the marker-stripped `control` fixture and the real `Excel_file_with_trash_item.xlsx`. `cg-<leg>-<fixture>-<n>.{inclusive,self}.txt` plus each run's stderr, and `summary.txt`, the differenced per-symbol table the record quotes. |
| `differential/oracle-summary.md` | The three differential instruments: the whole-corpus transcript (row counts and sha256 per leg), the 326-part verdict census, and the 60-package first-error matrix. |
| `differential/verdicts-before.tsv`, `verdicts-after.tsv` | Per worksheet part: the scan verdict, the bytes pulled from the part reader, and the part size. 326 parts from 180 packages, both legs. |
| `differential/matrix-before.tsv`, `matrix-after.tsv` | The 0541-style first-error matrix over 60 synthetic packages read through the public `SourceWorksheet::cell` and `cells`, both legs, 900 rows each. |
| `differential/matrix-witnesses.txt` | Every distinct (package, before, after) triple the matrix moves: the complete witness list behind the record's table. |
| `differential/make_matrix.py` | The generator for those 60 packages — change 0597's script, retargeted to this repository's tracked `page_scale.xlsx`. Deterministic; the packages themselves are not retained. |
| `probe/` | The scratch probe used for the instruction counts, the transcripts, the verdict census and the ineligible-read timing (`main.rs`, `Cargo.toml` with path dependencies). Modes: `ir`, `time`, `oracle`, `eagoracle`, `verdict`, `source`, `eager`. It is change 0597's probe plus the `time` and `verdict` modes. |
| `run_callgrind.sh`, `run_timing.sh`, `analyse_timing.py`, `analyse_probe_timing.py` | The capture and analysis scripts, as run. |
| `timing/` | The ineligible-read paired timing (`probe-<fixture>-{A1,B1,B2,A2,F1..F4}.txt`, one elapsed-ns line per sample) and the harness's eligible controls (`{A1,B1,B2,A2,F1..F4}.json`), plus `probe-summary.txt` and `harness-summary.txt`. |
| `gates.txt` | The tail of every gate that was run. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`, for the coordinator to merge. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `cleanup.json` | What was removed from the session scratchpad and the worktree area. |

## Provenance

- Base commit: `70d7768cc` (change 0652, the owner's decision record).
- Branch: `perf/0658-xlsx-selected-cell-ineligibility-gate`.
- Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws. Eight agents
  were building concurrently throughout; every measured process was pinned with
  `taskset -c 13`.
- Toolchain: `rustc 1.95.0`, `valgrind-3.26.0`.
- Binaries, all staged outside their Cargo target directories before any leg
  ran (sha256):
  - probe, before leg `8e72073b0a8454d9393c1070e8b554b0ae3d09948be956098588fa471a9ed378`
  - probe, after leg `829ebe31e3550cdc391ed445c6c325bbdf36a40d4aa3dd774b65275a0888b146`
  - `litchi-perf-baseline`, before leg `ba5b88af706e7a80b5dd870c77527dbaac9bfc513cf0be91453db61b9c5d9f24`
  - `litchi-perf-baseline`, after leg `24d539173bc000d61866286b654accc17a9d95486494277b63dc27e4db32d302`
  - The before legs were built from the shared read-only checkout
    `litchi-worktrees/before-70d7768cc` with their own `CARGO_TARGET_DIR`;
    `lto = true` makes release binaries non-reproducible byte for byte (0635),
    which is why these hashes are recorded.
- Fixtures: `test-data/poi/test-data/spreadsheet/Excel_file_with_trash_item.xlsx`
  (tracked) and change 0587's marker-stripped derivative of it, retained at
  [`../change-0587/xml-substrate/control.xlsx`](../change-0587/xml-substrate/control.xlsx)
  and not copied again here.

## What is not in this packet

The two 4,786-line corpus transcripts (about 1 MB each) and the 60 generated
synthetic packages. `differential/oracle-summary.md` carries the transcripts'
sha256 sums and `probe/` plus `differential/make_matrix.py` regenerate both
deterministically from tracked fixtures.
