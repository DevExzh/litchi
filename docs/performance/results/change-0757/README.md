# Change 0757 evidence packet

Record: [0757-xls-fresh-writer-sst-determinism](../../0757-xls-fresh-writer-sst-determinism.md).
Base `9ff78bbf1c` (head of `perf/0753-legacy-fresh-writer-text-paths`); measured
code `921787e2c0`; final head `a400341bdf` (branch
`perf/0757-xls-fresh-writer-sst-determinism`). `performance_claim: none`.

## Contents

| path | what it is |
| --- | --- |
| `environment.txt` | host, kernel, toolchains, pinning, measurement windows |
| `binaries.sha256` | SHA-256 of the measured harness (`lpb`) and probe (`prb`) binaries per leg, and of the final-head harness used only for the instruction check |
| `gates.txt` | every gate command at the final head with its exit code and test counts, the base-side clippy and golden runs, and the final-head instruction check |
| `latency/window-1/`, `latency/window-2/` | ABBA runs of the harness selectors (`scripts/abba_harness.py`): `summary-xls.json`, `summary-docppt.json` (per-process p50/p95/mean, paired ratios, bootstrap CI, output SHA-256 per leg) and every raw harness report in `raw/` |
| `counters/` | differenced `perf stat` counters (`scripts/perf_counters.py`): `summary.json` and the raw `perf stat -x,` outputs and harness reports in `raw/` |
| `probe/` | `write_to`-only probe runs, eight processes per case (`scripts/probe_runs.py`): `summary.json` (timings, callgrind instructions per write, allocation counts, output length and FNV-1a per process), raw JSON per process, callgrind logs |
| `probe-multi-default/`, `probe-multi-pinned/` | the two multi-string probe cases with sixteen processes, glibc's default malloc thresholds and with `trim_threshold`/`mmap_threshold` pinned at 256 MiB |
| `tunables/` | `scripts/tunables_check.py`: `xls_fresh_write_to/large` and `ppt_fresh_write_to/payload-heavy` under default and pinned malloc thresholds (page faults, user instructions, timed p50) |
| `ppt-long/` | `ppt_fresh_write_to/payload-heavy` with 200 samples per process (steady state against the first 40) |
| `behaviour/` | what each leg does with strings past each limit (`base-probe.txt`, `after-probe.txt`, from `probe-src/string_fields_probe.rs`); the golden test file run on the base before and after its new digest was pinned (`base-goldens-run.txt`, `base-goldens-final.txt`); the data-validation and empty-number-format findings (`findings-probe.txt`) |
| `probe-src/` | the timing/allocation probe (`main.rs`, `Cargo.toml.template`) and the string-field behaviour probe (a test file copied into `crates/litchi-xls/tests/` of a leg, run, and removed) |
| `build-logs/` | the harness build logs of both legs and of the final head |
| `scripts/` | the build, measurement and table scripts used |
| `tables.md` | the tables rendered from this packet by `scripts/tables.py` |
| `log-sections.md` | ready-to-paste paragraphs for `HOTSPOTS.md`, `REPORT.md` and `GOAL_AUDIT.md` |
| `cleanup.json` | what was removed after the evidence was copied here |

## Reproducing

1. Build each leg with `scripts/build_leg.sh SRC TARGET LOG` (identical command:
   `cargo build --release --locked --offline --manifest-path
   tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline`), before from a
   detached worktree at `9ff78bbf1c`, after from the branch. Copy the binaries to
   equal-length paths (`bin/A/lpb`, `bin/B/lpb`) under the scratch directory the
   scripts name.
2. `python3 scripts/abba_harness.py OUT 16 selector/shape,... LABEL` for
   timings; `python3 scripts/perf_counters.py OUT 16` for counters;
   `python3 scripts/tunables_check.py OUT 16 CASE SHAPE` for the heap check.
3. Fill `probe-src/Cargo.toml.template` with each leg's source tree, build it
   with `scripts/build_probe.sh` (`cargo build --release --offline`, seeded with
   the workspace `Cargo.lock`), copy to `bin/A/prb` and `bin/B/prb`, then
   `python3 scripts/probe_runs.py OUT 16 [cases [rounds [pinned]]]`.
4. `python3 scripts/tables.py .` renders `tables.md` from this packet.

The multi-string golden digests in `crates/litchi-xls/tests/xls_writer_text_goldens.rs`
were recorded from the branch (no earlier writer produced them reproducibly);
the `one_string_per_sheet` digest was recorded from the base with the at-limit
strings, and the final test file passes its pre-0753 test on the base
(`behaviour/base-goldens-final.txt`).
