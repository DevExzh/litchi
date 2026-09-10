# XLSB drawing-skip projection performance evidence

This directory contains the final release profile for the opt-in XLSB drawing
projection workload. It measures one checked-in Apache POI fixture and records
raw per-process JSON, source and binary provenance, and the exact replay
scripts. The result is scoped evidence for the named fixture and cases; it is
not a native Excel, full CRUD, allocation, RSS, or package-wide performance
claim.

The fixture is `test-data/poi/test-data/spreadsheet/testVarious.xlsb` (22,715
bytes, SHA-256
`8c600e97d719b0266dcfb49c1872feb8d10c6ed12bc768ff16ace7dae555ebfc`). The
release binary was built with Rust 1.95.0, `CARGO_INCREMENTAL=0`, two Cargo
jobs, and `CARGO_PROFILE_RELEASE_DEBUG=0`. Its SHA-256 is
`07f44c8b63fd46b94b66d024871f15b1933cea04062ad47134638ff1e54bdbca` and its
size is 11,736,696 bytes. The locked harness input is
`tools/perf-baseline/Cargo.lock` (SHA-256
`e1c66239cd25a5657a72ed48ece0cd39938a734867f621eeeacdf1217b04d14b`);
`Cargo.lock` is also recorded because the repository has two workspaces.
The provenance records the host as Linux 7.0.0-1011-aws on an AMD EPYC 9R45
with 32 logical CPUs and 132,553,797,632 bytes of reported memory.

The matrix has 26 supported backend/case combinations, three fresh processes
per combination, three warmups, and 30 timed samples per process. The raw run
therefore has 78 JSON reports in
[`runs/final-release`](runs/final-release). The lanes are:

| Backend | Cases | Scope |
| --- | ---: | --- |
| `owned` | 8 | Facade identification/open plus the default eager facade/direct operations, including facade-only `full_text` |
| `owned_direct` | 7 | Direct eager XLSB owner operations |
| `owned_without_drawings` | 7 | Direct owner with typed drawing parsing skipped; OPC/raw drawing members and relationships remain preserved |
| `source_backed` | 4 | Read-only source-backed open/catalog/selected-cell/full-scan operations |

Each JSON report contains its timing scope, p50/p95/p99 and all 30 raw sample
durations, correctness gates, corpus identity, and (for `source_backed`) the
instrumented positional-read and cache observations. The source-backed lane
does not run the eager cell-limit refusal probe, so
`tight_cell_limits_refused=false` there is an unavailable scope marker. It is
not a failed semantic gate.

`matrix-summary.json` is generated from the raw reports by
[`verify-matrix.py`](verify-matrix.py). It checks the 26×3 matrix, one binary
and fixture identity, 30 samples per process, all applicable gates, and
source-counter stability. It summarizes repeated-process statistics within
each backend/case only. Timing values are descriptive within each lane;
backend API and validation scopes differ, so no equivalent-work speedup is
inferred. Source-backed timing also includes the `ReadAt` counter and cache
instrumentation.

[`performance-receipt.json`](performance-receipt.json) records the final
binary, lockfile, source-stability, summary, and raw-JSON manifest hashes. On a
fresh checkout, the optional provenance replay also needs the repository
workspace lock copy: from the repository root, run `cp
docs/report/spec-gap-validation-evidence/xlsb-drawing-projection/gates/workspace-Cargo.lock
Cargo.lock` before running `record-provenance.py`; the release build itself
uses the tracked `tools/perf-baseline/Cargo.lock`.

The build and raw-run commands are retained in
[`release-build-command.txt`](release-build-command.txt),
[`run-release-matrix.sh`](run-release-matrix.sh), and
[`runs/final-release/commands.txt`](runs/final-release/commands.txt). To replay
in a detached checkout, first ensure the fixture path and the locked harness
workspace are available, then run:

```sh
export CARGO_TARGET_DIR=/var/tmp/litchi-xlsb-drawing-projection-replay-target
export CARGO_INCREMENTAL=0
export CARGO_BUILD_JOBS=2
export CARGO_PROFILE_RELEASE_DEBUG=0
cargo +1.95.0 build --manifest-path tools/perf-baseline/Cargo.toml \
  --offline --locked --release --features xlsb-crud --bin xlsb_crud
export XLSB_CRUD_BINARY="$CARGO_TARGET_DIR/release/xlsb_crud"
export XLSB_CRUD_RUN_ROOT="$PWD/docs/report/spec-gap-validation-evidence/xlsb-drawing-projection/performance/runs/replay"
export XLSB_CRUD_WARMUP=3
export XLSB_CRUD_SAMPLES=30
docs/report/spec-gap-validation-evidence/xlsb-drawing-projection/performance/run-release-matrix.sh
python3 docs/report/spec-gap-validation-evidence/xlsb-drawing-projection/performance/verify-matrix.py "$XLSB_CRUD_RUN_ROOT"
```

The provenance recorder writes the exact source, both lockfile, fixture,
toolchain, and binary hashes used by the release receipt. The post-build
verification found source and fixture hashes unchanged across the final build;
the before and after records are
[`source-provenance-before.json`](source-provenance-before.json) and
[`source-provenance-after.json`](source-provenance-after.json).
