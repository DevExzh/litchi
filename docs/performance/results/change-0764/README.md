# Evidence for record 0764 (XML attribute DoS hardening)

Record: [../../0764-xml-attribute-dos-hardening.md](../../0764-xml-attribute-dos-hardening.md).
Base `1d1044e3ac`; branch `perf/0764-xml-attribute-dos-hardening`.

| path | what it holds |
| --- | --- |
| `binaries.sha256` | SHA-256 of the four measured binaries. `bin/a` is the before leg: a detached worktree of the base with only the harness commits (`5d8db7ad29`, `b86c7567bc`, `3a04167848`) applied, built with the identical command as the after leg (`cargo build --release --locked --offline --manifest-path tools/perf-baseline/Cargo.toml --bin xml_attribute_bounds --bin litchi-perf-baseline`, `CARGO_BUILD_JOBS=6`; the after leg also set `CARGO_INCREMENTAL=0`, which release builds use anyway). Both copied to paths of equal length before running. |
| `adversarial/<case>/` | ABBA raw reports (`<case>-r<round>-<leg>.json`), `perf stat` counters (`.perf.csv`) and `<case>-summary.json` for the hostile inputs of `xml_attribute_bounds adversarial`. |
| `controls/<case>/` | The same for the benign controls: real-part MCE, stream, audit and style-reader cases of `xml_attribute_bounds`, and five `litchi-perf-baseline` cases. |
| `differential/` | Summary of the before/after differential (`compare.json`) and the SHA-256 of the two full reports, which are 44 MB each and not kept. |
| `census/` | `attribute_census.py` and its summary over the repository's real packages. |
| `survey/` | The read-only classification of every `.attributes()` call site in the OOXML crates (five tables), which chose the sites fixed. |
| `site-timings/` | W1's debug-build timings of its new hostile-input tests against the base code and against the fix; W4's malformed-input probes on the base code and on the fix. |
| `quick-xml-probe/` | A 70-line program against quick-xml 0.41 showing its recovery after a duplicate attribute and the cost of a lenient reader on repeated names; `output.txt` is its output on this host. |
| `scripts/` | `abba.py` (one ABBA series), `run_abba.sh` (the campaign), `compare_differential.py`. |
| `review/` | The review follow-up: `long-uri/` (the harness's long-URI cases and the MCE benign controls, before `5aac78a870` and after `2a5d896000`, one process per arm, with `summary.txt`), `review-inputs/` (the review's probe inputs, two ABBA rounds, with `perf stat` counters and `summary.txt`), `benign-regression/` (the first identity build's worksheet regression and its removal: counters, callgrind `memcmp` callers, `notes.txt`), `differential-compare.json` and `differential-full-reports.sha256` (the before/after differential, byte-identical), `binaries.sha256`, `gates.txt` and `tests.txt` (per test binary). |
| `gates.txt` | Gate commands and exit codes. |
| `cleanup.json` | What was removed after the evidence was copied here. |
| `log-sections.md` | Paragraphs for HOTSPOTS.md, REPORT.md and GOAL_AUDIT.md. |

Method: 4 ABBA rounds (A1 B1 B2 A2), each process pinned with `taskset -c 16`
under `perf stat -e instructions:u,cycles:u`; 15 retained samples after 3
warmups for hostile inputs and main-harness controls, 31 after 3 for the
real-part controls. Per process: p50, p95, mean. Per case: the median of
process p50s per arm, the paired after/before ratios of process p50s (eight
pairs), and a percentile bootstrap interval of the median ratio (10,000
resamples, seed 764). Instruction and cycle counts cover the whole process,
input construction included. Other agents were building and measuring on the
same host.
