# Change 0409 evidence

This is descriptive XLSX evidence, not an accepted before/after claim.

- `initial/`: source abe38a9570129c6646bb1b1d7207c407fc86c3d6; seven reports and their exact command journal, CPU profile, flamegraph, catalogs and hashes. The source-edit workbook/worksheet counters are unconfigured false zeroes; total source, timing, sink and preservation evidence remain valid.
- `corrected/`: committed harness range correction; source-edit 500 samples plus two-sheet, all-sheet batch, and managed smoke checks. `compressed-member-intersections-v1` identifies cumulative compressed source ranges across open/planning/commit/publication. Positive unselected reads include untouched raw publication, not semantic decoding.
- `attribution/`: reproducible period-weighted stack parser and the all-symbol/no-inline profile; `SelectedWorksheet::cell` identifies a subset of the timed query. Inclusive rows overlap. Sample periods do not measure wall-clock phase duration.
- `pmu/`: local sysfs/perf inventory and controlled native-L2 event probes. Exact LLC events are unavailable and all-zero generic L1 aliases are unusable on this guest. Each probe event was measured in a separate invocation.
- `checks/`: build/test/validation logs, including initial formatting and pinned-toolchain failures with their corrected checks.

The command journals retain original temporary paths. Reproduction requires rebuilding the recorded source with Rust 1.98.1 and the exact recorded environment, then adjusting the task-specific output paths. The capture scripts assert the original source/worktree and deterministic corpus identities. The user's untracked docs/GOAL.md is recorded as the sole dirty entry. No clean ABBA, cold-cache, scaling, full CRUD coverage or production speedup claim is authorized.

Large symbolized perf scripts are retained losslessly as Zstandard. Raw perf.data, symbolized reports, the flamegraph, and binaries' hashes/build IDs remain; the large build executables and PMU probe executable are regenerated from the retained sources/commands. Original capture artifacts.json describes the original temporary capture, including the now-compressed script. The top-level artifact-manifest.json is authoritative for the published layout.

Run `python3 verify-artifacts.py` from any directory to verify every published file and both decompressed scripts. Zstandard is required. The parser in attribution/ can consume a decompressed script to reproduce its JSON summary.
