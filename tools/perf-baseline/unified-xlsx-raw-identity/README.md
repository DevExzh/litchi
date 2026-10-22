# XLSX raw identity lifecycle probe

This probe compares the owned timeline raw insertion lifecycle at two exact
source revisions. It covers a one-cache small package, a one-cache package
with sixteen deterministic 512 KiB parts, an eight-cache batch on the small
package, and the same batch on the multipart package. Each case constructs
the package and raw values before the measured region, performs
`Package::edit_timelines` → `timeline::raw::Cache` batch insertion → `commit`,
and verifies the output bytes and native relationship IDs afterward.

The probe was added in `44c6d22be72a6351e48fb92e4c18017101657bcb`. To measure
the pre-repair source, check out `47c89d20da1594187c967e161b506d828aed1599`
and cherry-pick that harness commit. Keep each target directory outside the
checkout and record the resulting binary SHA-256 in the report:

```sh
# Baseline checkout: 47c89d20da1594187c967e161b506d828aed1599 plus the
# harness commit above.
CARGO_NET_OFFLINE=true cargo build --release --locked \
  --features allocator-metrics \
  --manifest-path tools/perf-baseline/Cargo.toml \
  --bin litchi-perf-baseline-alloc \
  --target-dir /outside/uo2-raw-baseline-target

taskset -c 0 /outside/uo2-raw-baseline-target/release/litchi-perf-baseline-alloc \
  xlsx-raw-identity \
  --api-variant original \
  --source-revision 47c89d20da1594187c967e161b506d828aed1599 \
  --iterations 5 --warmups 2 --json /outside/report-original.json

# Candidate checkout: 44c6d22be72a6351e48fb92e4c18017101657bcb.
CARGO_NET_OFFLINE=true cargo build --release --locked \
  --features allocator-metrics,xlsx-raw-candidate \
  --manifest-path tools/perf-baseline/Cargo.toml \
  --bin litchi-perf-baseline-alloc \
  --target-dir /outside/uo2-raw-candidate-target

taskset -c 0 /outside/uo2-raw-candidate-target/release/litchi-perf-baseline-alloc \
  xlsx-raw-identity \
  --api-variant candidate \
  --source-revision 44c6d22be72a6351e48fb92e4c18017101657bcb \
  --iterations 5 --warmups 2 --json /outside/report-candidate.json
```

The baseline call uses the old required-limits method through the feature
adapter; the candidate call uses the retained edit policy. For instruction
counts, run the same command for each variant and case under Callgrind, using
the same iterations, warmups, CPU pinning, and output directory outside the
repository:

```sh
valgrind --tool=callgrind --callgrind-out-file=/outside/raw-baseline-small.out \
  /outside/uo2-raw-baseline-target/release/litchi-perf-baseline-alloc \
  xlsx-raw-identity --api-variant original \
  --source-revision 47c89d20da1594187c967e161b506d828aed1599 \
  --case small_single --iterations 1 --warmups 1

valgrind --tool=callgrind --callgrind-out-file=/outside/raw-candidate-small.out \
  /outside/uo2-raw-candidate-target/release/litchi-perf-baseline-alloc \
  xlsx-raw-identity --api-variant candidate \
  --source-revision 44c6d22be72a6351e48fb92e4c18017101657bcb \
  --case small_single --iterations 1 --warmups 1
```

The report records the generated corpus SHA-256 for every case, the source
revision supplied by the operator, binary identity, output SHA-256, allocation
samples, and the measured-region scope. It does not include output
serialization, identity verification, or destruction in the allocation
region. The Callgrind command supplies the separate Ir measurement; its
capture and summary files stay outside the repository.
