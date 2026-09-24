# Evidence packet for change 0763

Record: [0763](../../0763-rough-budget-leases.md) — rough budget leases
(`Budget::lease`, `ExecutionContext::lease`) for the per-paragraph and per-row
charges of the streaming DOCX and XLSX writers. Base `6ec785c265` (the tip of
record 0762 on the same branch), candidate `1ab98ca669`. Every measured
process was pinned to CPU 12 on an AMD EPYC 9R45.

| path | what it holds |
| --- | --- |
| `binaries.txt` | build commands (identical for both legs) and SHA-256 of every measured binary; the probes |
| `gates.txt` | every gate command, exit code and per-crate test count |
| `timing/abba-raw.tar.gz` | the 128 raw harness reports of the ABBA campaign (4 rounds × 8 cases × before, after, after, before), with `progress.log` and `stderr.log` |
| `timing/summary.json`, `timing/flags.json` | per-case medians, paired changes, bootstrap intervals, per-leg output digests, per-process rows; every paired p50/p95/mean comparison beyond 5% |
| `counters/` | per-timed-iteration user-space instructions and cycles (low/high-sample differencing, two ABBA rounds), raw and summary |
| `alloc/` | allocation-build reports (one process per leg and case) and the per-case medians |
| `probe/lease_diff/` | the base-versus-candidate differential probe: source and manifest template (`ROOT` is each tree) |
| `probe/lease-diff-summary.txt` | both transcripts' line counts and SHA-256 (identical), their first lines and the outcome categories |
| `probe/atomic_count/`, `probe/atomic-count.txt` | the budget-update count probe and its output for the large DOCX and XLSX scripts |
| `profiles/` | flat `perf report` summaries (whole process) of the large DOCX and XLSX cases, before and after |
| `scripts/` | every runner and summarizer used (shared with change 0762) |
| `cleanup.json` | what was removed and what was kept, for both records on this branch |
| `log-sections.md` | ready-to-paste sections for `HOTSPOTS.md`, `REPORT.md` and `GOAL_AUDIT.md` |

## Replaying

Build each leg with the commands in `binaries.txt`, copy the binaries to
`bin/b/` and `bin/a/` of equal-length paths, then run `scripts/run_abba.sh`,
`scripts/run_counters.sh` and `scripts/run_alloc.sh` with their summarizers
(the scripts carry this host's absolute paths). `lease_diff` builds twice, once
per tree, after replacing `ROOT` in its manifest template and copying the
workspace `Cargo.lock` in; run both binaries and compare stdout byte for byte.
`atomic_count` builds against one tree and prints its counts.
