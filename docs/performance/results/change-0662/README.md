# Retained evidence — change 0662

Record: [`../../0662-parallel-changed-member-deflate.md`](../../0662-parallel-changed-member-deflate.md).
Authority: change [0652](../../0652-owner-decisions-for-the-third-wave.md),
decision 6, which accepts ADR 0031 for bounded parallel deflate of changed
members. The design boundary and the prerequisite measurements are from
[0624](../../0624-parallel-changed-member-deflate-design.md), with the
execution-context shape from [0615](../../0615-execution-context-completeness-design.md).

This packet retains the implementation evidence for the preservation writer's
changed-member deflate wave. It does not claim that the other three ADR 0031
session rows are complete: ZIP bulk reads, CFB bulk reads and source-backed
multi-Part reads still need their own budget-composition implementation and
verification. The record calls that limitation out explicitly.

## Provenance

| | |
| --- | --- |
| base commit | `ab07e2a47` (change 0654) |
| branch | `perf/0662-parallel-changed-member-deflate` |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| toolchain | `rustc 1.95.0` |
| build | workspace release profile with LTO and `panic = "abort"`; targeted gates use `--locked`; probes use their retained manifests |
| scheduling | local Rayon pool is lazy and private; caller facilities are explicit; no global pool is initialized |
| concurrent work | other agents were active during the retained measurements; every reported table includes its same-window A/A floor |

## Contents

| path | what it is |
| --- | --- |
| [`scaling/zip-sweep.tsv`](scaling/zip-sweep.tsv) | warm-session ZIP publication sweep for the remainder balance rule |
| [`scaling/zip-sweep-tiny.tsv`](scaling/zip-sweep-tiny.tsv) | small-shape crossover sweep |
| [`scaling/opc-sweep.tsv`](scaling/opc-sweep.tsv) | cold-session and caller-facility OPC publication sweep |
| [`scaling/contended-window.tsv`](scaling/contended-window.tsv) | retained saturated-host counter-example |
| [`timing/fixture-timing.tsv`](timing/fixture-timing.tsv) | real-fixture end-to-end publication timing at widths 1/2/4/8 |
| [`timing/harness-ab.tsv`](timing/harness-ab.tsv) | paired default-path before/after harness legs |
| [`timing/fixture-timing-cpuset.txt`](timing/fixture-timing-cpuset.txt) | CPU-set and run-window notes for the fixture timing |
| [`counts/changed-set-census.tsv`](counts/changed-set-census.tsv) | deterministic eligible changed-set census over OOXML fixtures |
| [`counts/faults.txt`](counts/faults.txt) | minor-fault and peak-RSS publication counts |
| [`identity/corpus-identity.txt`](identity/corpus-identity.txt) | byte-identity and refusal differential across the fixture corpus |
| [`probe/`](probe/) | source, manifest template and drivers for replaying the retained probes |
| [`gates.txt`](gates.txt) | targeted validation commands and their results |
| [`decision.json`](decision.json) | decision, evidence, costs, limitations and provenance |
| [`log-sections.md`](log-sections.md) | four paragraphs for the shared performance logs |

## Replay

The deterministic correctness tests run from the workspace root:

```sh
CARGO_BUILD_JOBS=2 cargo fmt --all --check
CARGO_BUILD_JOBS=2 cargo check -p litchi-core -p soapberry-zip -p litchi-opc --all-targets --locked
CARGO_BUILD_JOBS=2 cargo test -p litchi-core -p soapberry-zip -p litchi-opc --all-targets --locked
CARGO_BUILD_JOBS=2 cargo clippy -p litchi-core -p soapberry-zip -p litchi-opc --all-targets --locked -- -D warnings
```

The retained probe drivers under `probe/` reproduce the ZIP and OPC sweeps;
they should be run only after reserving a quiet CPU set because the scaling
results are sensitive to host contention. The real-fixture scenario replaces
the first 64 Deflate XML parts with deterministic compact payloads and compares
the published digest at every width and with a caller facility.

## Scope and limits

The implemented budget bridge covers `Resource::Workers`, `Resource::CpuTasks`,
the write-side memory reservation, cancellation and the explicit scoped-worker
facility for managed source-backed publications. It intentionally leaves
`Resource::IoConcurrency` and the three existing read sessions for a later
record; no read-session budget-composition result is included here. The
ordinary save, eager writer, unmanaged packages, and splice/replay route do not
carry an execution context and therefore do not take this write wave.

No performance claim is registered. The paired medians, A/A floors, corpus
counts and RSS figures are retained evidence scoped to this host, build,
fixtures and publication scenarios.
