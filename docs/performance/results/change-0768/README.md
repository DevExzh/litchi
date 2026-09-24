# Evidence packet for change 0768

Record: [0768-doc-protection-classification](../../0768-doc-protection-classification.md).
Base `1d1044e3ac`; measured branch head `ac63852fc3` (the three code commits).
The review follow-ups (`95b485a290`, `de2df97e1f`, `0792812830`, `83fa8563d7`)
were gated and tested but not re-measured; their tests are in the crate, and
their counts are in the record and in `gates.txt`.

## Contents

| path | what it is |
| --- | --- |
| `corpus-probe/ole-doc-base.jsonl`, `ole-doc-after-final.jsonl` | `scripts/probe_0768.rs` over the 38 files of `test-data/ole/doc`, base and branch: FIB/DOP shape, raw DOP protection bits, `Selsf` CPs, the public classification, and a tracked insertion at CP 0 and CP 1 (with strict reopen and byte-exact republish) |
| `corpus-probe/other-doc-*.jsonl` | the same probe over the other 19 `.doc` files in `test-data` |
| `corpus-probe/*-after-classifier-and-selsf.jsonl` | the same probe after the first two commits only: the three "malformed CHPX FKP" reopen failures that led to the third commit |
| `corpus-probe/ole-doc-table.md`, `other-doc-table.md` | the per-fixture tables (`scripts/fixture_table.py`) |
| `latency/batch1`, `latency/batch2` | two ABBA batches, eight rounds each (`scripts/abba.py`): `raw.tar.gz` holds every harness and probe JSON report and `perf stat` CSV, `schedule.json` the process order, argv, exit code and load average, `analysis.json`/`summary.txt` the output of `scripts/analyze.py` |
| `latency/default-large-only-first` | the first `--writer-shape large`-only observation (two sequential processes per leg) with page faults |
| `latency/default-large-only`, `latency/pinned-malloc` | four ABBA rounds of the `large` shape alone, with default glibc and with glibc's mmap/trim thresholds pinned (`scripts/large_only.sh`, `scripts/large_only_summary.py`) |
| `latency/superseded-toolchain-schedule.json` | the schedule of a first ABBA run discarded before analysis because its probe binaries had been built by the host's default toolchain (1.98.1) instead of the pinned 1.95.0; its raw reports were deleted |
| `counters/` | exact timed-region instruction counts from callgrind (`scripts/cg_extract.py`), the 30 largest inclusive differences of the harness `large` run, and the Vec-growth callers; the callgrind profiles themselves were deleted |
| `binaries.sha256` | the measured binaries, fixtures and the probe lockfile |
| `environment.txt`, `gates.txt`, `cleanup.json`, `log-sections.md` | host, gate commands and exit codes (the original change and, appended, the review follow-ups), removed directories, ready-to-paste log sections |
| `scripts/diag_0768.rs` | the diagnostic that located the overflowing CHPX FKP page |

## How the binaries were built

- Harness: `cargo build --release --manifest-path tools/perf-baseline/Cargo.toml
  --locked --offline --bin litchi-perf-baseline`, identically in the branch
  worktree and in a detached worktree of the base, each with its own
  `CARGO_TARGET_DIR`, `CARGO_BUILD_JOBS=6`, Rust 1.95.0.
- Probe: the 0730 public-lifecycle probe
  (`docs/performance/results/change-0730/probe`), unchanged in code. Its tracked
  lockfile no longer matches today's crates, so both legs were built from scratch
  copies whose path dependencies point at the respective worktree, sharing one
  lockfile generated offline from the workspace lockfile (hash in
  `binaries.sha256`): `cargo build --release --locked --offline --bin
  ole_format_save_probe`, `RUSTUP_TOOLCHAIN=1.95.0`.
- Every binary was copied to `bin/A/…` (base) or `bin/B/…` (branch) so argv[0]
  lengths match; all other arguments are identical across legs.

## Reproduction

With the two legs staged as above, from the staging directory:

```sh
python3 abba.py runs 8 6          # one batch; repeat into runs2
python3 analyze.py runs runs/analysis.json
sh large_only.sh                  # large shape alone, default and pinned malloc
valgrind --tool=callgrind --toggle-collect='*timed_format*' bin/A/probe \
  --case docnohf --input NoHeadFoot.doc --operation format --warmups 0 --samples 8
valgrind --tool=callgrind bin/A/harness --case doc_semantic_one_edit_save \
  --writer-shape large --samples 2 --warmup 0 --json /dev/null
python3 cg_extract.py cg callgrind-summary.json
```

The corpus probe runs as a temporary `litchi-doc` example:
`cargo run -p litchi-doc --example probe_0768 --locked --offline -- test-data/ole/doc`.
