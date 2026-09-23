# Change 0756 evidence packet — wave integration

Integration record: [0756](../../0756-ole2-ooxml-wave-integration.md).

## Contents

- `base-sweep/` — the coordinator's opening sweep of about 100 OLE2/OOXML harness
  cases on base `009d515bef` (`semantic.json.gz`, `xlsx.json.gz`,
  `misc.json.gz`; 7 samples, 2 warmups, core 20). `run.sh` is the command
  script as run. Its output paths point at a session scratchpad that no longer
  exists. The binary was the base harness built with
  `cargo build --release --offline --locked --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline`.
  This sweep chose the wave's targets and is descriptive only.
- `profile-r2/` — timed-region attribution of nine base cases:
  - `REPORT.md`;
  - the three streaming summaries;
  - `folded/`, gzipped period-weighted stacks plus the scripts that reproduce
    every figure.

  The raw `perf script` texts were deleted after folding.
- `sweep/` — the wave-wide before/after sweep:
  - `README.md`, with exact commands, binary SHA-256, environment and caveats;
  - `summary.md` and `summary.json`;
  - `raw/`, 48 gzipped reports plus `perf stat` output and the run log;
  - `docopen-check/`, the coordinator's instruction-count follow-up on
    `doc_semantic_open` (large).
- `gates/` — the final integration gate logs on the integration tip, the same 16
  gates as change 0675's runner, run with `CARGO_BUILD_JOBS=16`.
- `implementer-briefing.md` — the shared briefing every implementer read. It
  covers constraints, worktree and measurement method, gates and deliverables,
  and the lessons added during the wave: identical-command before legs,
  code-layout sensitivity, TMPDIR, and foreground cleanup.
- `log-sections.md` — this record's sections for the shared logs.
- `cleanup.json` — what was removed, kept and preserved.

## Replay

The sweep's summary is recomputed from `sweep/raw/` by the summary script
embedded in `sweep/README.md`. The gate commands are listed in each gate log's
header line.
