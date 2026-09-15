# Evidence: change 0597, the selected-cell ineligibility gate and the post-mark merge refusal

Change record:
[`0597-xlsx-selected-cell-ineligibility-gate.md`](../../0597-xlsx-selected-cell-ineligibility-gate.md).

Disposition: **mixed**. `performance_claim: none`. One correctness fix is
implemented in `crates/litchi-xlsx/src/raw/worksheet/selected.rs`; the
ineligibility gate that this change set out to build is **not implemented** and
is frozen as a design, with its measured prize and its rejected patch retained
here. No timing claim is made: every paired delta is inside its scenario's own
A/A floor, and that floor exceeds 24% at p50 on two of the seven scenario/shape
pairs measured in this window.

## Contents

| Path | What it is |
| --- | --- |
| `callgrind/` | Isolation-pair annotations (inclusive and self, N = 1 and N = 11) for the source-backed one-cell read on the marker-stripped `control` fixture and the real `Excel_file_with_trash_item.xlsx`, for three legs: `before` (base `08d968f8e`), `fix` (this change), `gate` (this change plus the frozen design candidate). `cg-<leg>-<fixture>-<n>.{inclusive,self}.txt`, plus each run's stderr. |
| `differential/oracle-summary.md` | The three-leg full-corpus transcript comparison: row counts, sha256 of each transcript, leg-to-leg differences, and what the defect fix moved. |
| `differential/oracle-changed-rows.tsv` | Every row the fix moves (434), in all three legs, with its logical read and byte counts and the exact `Result` debug. |
| `differential/eager-agreement.txt` | The same grid read through the fully materialized `litchi_xlsx::Workbook`: 372 base mismatches, 0 after. |
| `differential/matrix-before.tsv`, `matrix-after.tsv` | The 0541-style first-error matrix over 60 synthetic packages, base and after, through the public `SourceWorksheet::cell` and `cells`. |
| `differential/make_matrix.py` | The generator for those 60 packages: one small valid package with `xl/worksheets/sheet1.xml` rewritten, malformed at a chosen position, in a `<cols>`-bearing shape and a `<cols>`-free control shape. |
| `design-candidate/gate-candidate.patch` | The frozen ineligibility gate, as a diff against the landed `selected.rs`. Not applied. |
| `probe/` | The scratch probe used for the instruction counts and the differential transcripts (`main.rs`, `Cargo.toml` with path dependencies). Modes: `ir` (N repeats for an isolation pair), `oracle` (the source-backed transcript over a counting positional source), `eagoracle` (the same grid through the materialized `Workbook`). |
| `run_callgrind.sh`, `run_timing.sh`, `analyse_timing.py` | The capture and analysis scripts, as run. |
| `timing/` | The eight harness runs: `A1`, `B1`, `B2`, `A2` (the ABBA pair) and `F1`..`F4` (the A/A floor in the same window), plus `abba-summary.txt`. |
| `gates.txt` | The tail of every gate: `cargo fmt --all --check`, `cargo clippy -p litchi-xlsx --all-targets`, `cargo doc -p litchi-xlsx --no-deps`, `cargo test -p litchi-xlsx`. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`, for the coordinator to merge. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `cleanup.json` | What was removed from the session scratchpad and the worktree area. |

## Provenance

- Base commit: `08d968f8ec7db27cf1187d01911fd08b9d014d91` (change 0587).
- Branch: `perf/0597-xlsx-selected-cell-ineligibility-gate`.
- Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws. Eight agents
  were building concurrently throughout; every measured process was pinned with
  `taskset -c 20`.
- Toolchain: `rustc 1.95.0 (59807616e 2026-04-14)`, `valgrind-3.26.0`.
- Harness binaries, `cargo build --release --locked` of `tools/perf-baseline`:
  - before leg `434b28b5298218918418ab251ff4cce40b35579ab5abf6215e2ca9cb5d628ef9`
    (built from the shared read-only checkout `litchi-worktrees/before-08d968f8e`)
  - after leg `98688c0a9b59871ea2f7df4b91fa422ee2972487bbf773f253d86ebd3af6321f`
- Probe binaries, `cargo build --release` (`debug = 1`):
  - before leg `c3383baafc611fcbcc35dc1fe9d9ec2d25d8044a2160f1240bf6c292a80dbc4a`
  - after leg `d856e292c084867bc6519f874200096899dbd67fcd5388c0ea44e8828ea9a70e`
  - The `gate` callgrind leg was profiled from a third probe built against this
    worktree with `design-candidate/gate-candidate.patch` applied; that binary
    was replaced when the patch was reverted, so its hash is not retained.
- Fixtures: `test-data/poi/test-data/spreadsheet/Excel_file_with_trash_item.xlsx`
  (tracked) and change 0587's marker-stripped derivative of it, retained at
  `../change-0587/xml-substrate/control.xlsx` (not copied again here).

## What is not in this packet

The three full 4,606-row transcripts are about 1 MB each and are not retained;
`differential/oracle-summary.md` carries their sha256 sums, and every row the
change moves is retained in full in `differential/oracle-changed-rows.tsv`. The
60 synthetic `.xlsx` packages are not retained either; `make_matrix.py`
regenerates them deterministically from a tracked fixture.
