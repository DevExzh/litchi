# XLSB worksheet binary-index profile

This directory owns the opt-in profiling harness for the source-backed XLSB
worksheet binary-index cached-value API. The harness is the direct
`litchi-xlsb` example at
[`crates/litchi-xlsb/examples/binary_index_profile.rs`](../../../../../crates/litchi-xlsb/examples/binary_index_profile.rs).
It runs one case in one process, validates every returned value against a
complete eager worksheet oracle prepared before timing, counts logical
`ReadAt` requests and bytes, records `SourceCacheDiagnostics`, and records
operation-scoped `std::alloc::System` counters. Allocator-instrumented elapsed
time is retained as measurement evidence; it is not production latency.
The allocator's requested-live highwater is reset immediately before each
timed interval, so `allocation.peak_before` is that interval's live baseline
and `allocation.peak_after` is its requested-live highwater. These fields are
not RSS or an allocator-internal highwater. Warm reports also retain setup
allocation counters and cache gauges for index/materialization preparation plus
the declared warmups. Both lanes use the source-backed and binary-index default
finite limits; no custom limit override is part of this profile.

Build and smoke-check the harness without the umbrella facade:

```sh
RUSTUP_HOME=/tmp/litchi-spec-gap-rustup CARGO_INCREMENTAL=0 cargo +1.95.0 check -p litchi-xlsb --example binary_index_profile
RUSTUP_HOME=/tmp/litchi-spec-gap-rustup CARGO_INCREMENTAL=0 cargo +1.95.0 run -p litchi-xlsb --example binary_index_profile -- \
  --case indexed_cold --warmup 3 --samples 30
```

The bounded matrix script runs `testVarious.xlsb` and `62815.xlsb` through
`indexed_cold`, `materialize_cold`, `indexed_warm`, and `materialize_warm`,
using a fresh process for each report. With `PROFILE_BIN` unset, the script
always performs a fresh offline locked release build in the selected target
directory, even if an executable is already present. To use an externally
frozen binary, set both `PROFILE_BIN` and `PROFILE_BUILD_MANIFEST`; the latter
must be a retained source/build manifest whose hash is checked before and
after the run:

```sh
RUSTUP_HOME=/tmp/litchi-spec-gap-rustup \
  ./docs/report/spec-gap-validation-evidence/xlsb-binary-index/performance/run-profile.sh
```

Set `PROCESSES=3` to repeat every fixture/case in three independent fresh
processes; the verifier keeps those reports separate and summarizes their
within-case statistics. The default is one process per fixture/case for a
bounded smoke run.

After the source/test freeze, the requested final matrix adds the larger
generated control and uses the frozen Rust 1.95 release binary:

```sh
export RUSTUP_HOME=/tmp/litchi-spec-gap-rustup
export RUST_TOOLCHAIN=1.95.0
export CARGO_TARGET_DIR=/var/tmp/litchi-xlsb-binary-index-final-target
export CARGO_INCREMENTAL=0
export OUTPUT_DIR="$PWD/docs/report/spec-gap-validation-evidence/xlsb-binary-index/performance/raw/final-release"
export PROCESSES=3 WARMUP=3 SAMPLES=30 SYNTHETIC_CELLS=131072
export PROFILE_BIN=/var/tmp/litchi-xlsb-binary-index-final-target/release/examples/binary_index_profile
export PROFILE_BUILD_MANIFEST="$OUTPUT_DIR/release-build-manifest.txt"
./docs/report/spec-gap-validation-evidence/xlsb-binary-index/performance/run-profile.sh
```

The retained run used the external binary and manifest shown above. For a
detached checkout after that temporary target is removed, unset both
`PROFILE_BIN` and `PROFILE_BUILD_MANIFEST`; the runner then performs the
offline locked release build itself and records a new build manifest.

The frozen final run produced 36 reports (three fresh processes for each of
three fixtures and four cases). The table reports the nearest-rank percentile
summary as the median of the three process summaries. `reads/bytes` are the
per-sample operation deltas; `op alloc` is the median requested allocation
across the 90 raw samples in that fixture/case. Warm `setup alloc/peak` is the
retained-handle preparation counter reported by the harness. Elapsed values
are allocator-instrumented nanoseconds and describe the declared case scopes;
they do not establish an equivalent-work production speedup.

| Fixture | Case | p50 / p95 / p99 (ns) | reads / bytes | op alloc (bytes) | warm setup alloc / peak (bytes) | retained cache bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `testVarious.xlsb` | `indexed_cold` | 117111 / 133801 / 138990 | 9 / 1433 | 1096894 | — | 4349 |
| `testVarious.xlsb` | `materialize_cold` | 136491 / 158341 / 158721 | 12 / 2173 | 1227095 | — | 6673 |
| `testVarious.xlsb` | `indexed_warm` | 340 / 390 / 490 | 0 / 0 | 21 | 259572 / 86132 | 4349 |
| `testVarious.xlsb` | `materialize_warm` | 190 / 200 / 410 | 0 / 0 | 21 | 389773 / 123339 | 6673 |
| `62815.xlsb` | `indexed_cold` | 114271 / 132950 / 137371 | 6 / 2641 | 868943 | — | 25126 |
| `62815.xlsb` | `materialize_cold` | 332211 / 348581 / 356201 | 6 / 2620 | 1349201 | — | 23990 |
| `62815.xlsb` | `indexed_warm` | 260 / 320 / 390 | 0 / 0 | 0 | 267030 / 105212 | 25126 |
| `62815.xlsb` | `materialize_warm` | 100 / 110 / 220 | 0 / 0 | 0 | 747288 / 245846 | 23990 |
| `synthetic:131072` | `indexed_cold` | 3086792 / 3101591 / 3109552 | 20 / 352138 | 4296310 | — | 2500095 |
| `synthetic:131072` | `materialize_cold` | 18002467 / 18141118 / 18185048 | 20 / 351122 | 71714427 | — | 2470736 |
| `synthetic:131072` | `indexed_warm` | 650 / 710 / 970 | 0 / 0 | 0 | 3698803 / 3277905 | 2500095 |
| `synthetic:131072` | `materialize_warm` | 130 / 140 / 520 | 0 / 0 | 0 | 71116920 / 70694189 | 2470736 |

The source observations were stable across all three processes. The synthetic
fixture's preparation counters show the retained binary-index path recording
about 3.7 MB of requested setup allocation versus about 71.1 MB for complete
worksheet materialization; these are `std::alloc::System` requested-byte
counters, not RSS or an allocator-internal highwater. The logical read and
cache counters are in-memory source-reader observations, and the compressed
worksheet payload is still materialized by the OPC layer in both cold cases.

The retained final raw reports, source-state manifests, external build
manifest, release build log, and binary hash are under
[`raw/final-release/`](raw/final-release/). The aggregate verifier output is
[`raw/matrix-summary.json`](raw/matrix-summary.json). The release binary used
for this run has SHA-256
`316d0d0d90ebec49978c9a9a08b6a60db82fd3bbb0d3b97d7ec194b49b3f5490`.

The runner records and checks the release build log (for an in-script build),
the manifest (for an external binary), the binary hash, and source-state
hashes alongside the provenance record.

For a larger allocation-shape control, set `SYNTHETIC_CELLS` to a bounded
positive count. The harness then uses the existing public XLSB writer to build
one deterministic 32-column worksheet before timing; generation is outside the
timed source-backed operation. For example, `SYNTHETIC_CELLS=131072` adds a
131,072-cell fixture to the same four cases. Its provenance records the
generation specification rather than a filesystem hash.

`indexed_cold` includes source-backed open, index/worksheet preparation,
cached-value reads, and source-backed handle teardown while the caller-owned
counting source remains held for a balanced allocation snapshot.
`materialize_cold` includes the same source-backed open, complete
selected-worksheet materialization, equivalent cached-value checks, and
teardown. Warm cases retain their prepared handle and time only repeated lookup;
their `setup_cache` and `setup_reads` fields include preparation and warmups.
The worksheet compressed payload is still materialized by the OPC layer in
both cold cases. Logical read counters therefore describe source-reader work,
not physical I/O or partial ZIP decompression. Retained cache bytes are cache
gauges, not heap or RSS measurements. The final matrix must be run only after
the source and test freeze; raw JSON, provenance, and the generated
`matrix-summary.json` are retained under `raw/`.
[`verify-matrix.py`](verify-matrix.py) recomputes nearest-rank percentiles and
the mean from raw samples, checks semantic/digest and source observations,
verifies interval allocator accounting, and summarizes each fixture/case
without comparing unlike work. `source_observation_stable` compares each
sample's operation read/cache deltas and retained gauges; cumulative counters
in the raw sample records are reported for context and are not used as a
stability gate.
