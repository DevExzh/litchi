# Change 0192: ODT open-only rerun evidence closure

Date: 2026-08-18

## Decision

Repeat only the withheld open-only workload of
[change 0191](0191-odt-unified-source-ingress.md) on clean current HEAD,
following the change 0183 precedent. No production or harness mechanism
changed. The rerun uses a release binary that is bit-identical to the change
0191 binary (SHA-256
`981ed4fbea8625b7d3feb4721262d992400bd522c51ce4be7b41071447129e59`), built
from current HEAD `9b717bb08` (change 0191) with unrelated pre-existing local
modifications outside the dependency graph.

This is evidence closure for the existing source-backed ODT filesystem open
path, not a new optimization and not a broader ODF claim.

## Measurement contract

The contract is unchanged from change 0191: fresh CPU-2-pinned processes in
`A1 eager, B1 source-backed, B2 source-backed, A2 eager` order, 30 warmups
and 500 retained samples per leg, one fresh process per leg, warm in-process
samples, uncontrolled page cache. The open-only control times
`fs::read(Path) + Document::from_bytes`; the source-backed path times
`Document::open(Path)` on the deterministic 16,812,034-byte media-rich ODT
(SHA-256 `29d8c1dcd21e739b07d95e463a875126af73d040188433e77817b47717e42bae`).

The predeclared p50/mean/p95/p99 same-implementation drift ceilings are
5%/5%/10%/15%. A statistic is accepted only when both paired directions are
lower and both implementation drifts pass its ceiling. Percentiles use the
`tools/perf_compare.py` method (p50 median, p95/p99 nearest-rank); drift is
`(second leg - first leg) / first leg`. Exactly one rerun was performed;
statistics that miss their gates remain withheld rather than retried.

## Result

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Eager drift | Source drift | Decision |
|---|---:|---:|---:|---:|---|
| p50 | 51.68% | 52.62% | 3.13% | 1.11% | accept |
| mean | 51.27% | 53.64% | 5.10% | -0.01% | reject: eager drift |
| p95 | 49.01% | 57.86% | 11.58% | -7.78% | reject: eager drift |
| p99 | 51.83% | 59.56% | 10.05% | -7.60% | accept |

Open p50 values are 3.458 ms -> 1.671 ms and 3.566 ms -> 1.690 ms; p99 values
are 4.391 ms -> 2.115 ms and 4.832 ms -> 1.954 ms. Both accepted statistics
improve by roughly half in every paired direction with drifts inside their
ceilings, so the change 0191 open-only path now has accepted warm latency
evidence at p50 and p99 for this exact corpus.

Open-only mean misses its 5% eager-drift ceiling at 5.10% and p95 misses its
10% eager-drift ceiling at 11.58%, so both remain withheld. The withheld 0191
run and this accepted rerun together indicate sub-millisecond leg-to-leg
eager drift on this host, not a candidate regression: every paired reduction
is near 50% in both directions.

Every leg passes the complete untimed gates: semantic parity, archive
identity, media identity, source-range coverage, and zero picture-payload
overlap. The independent instrumented replay records exactly 30 source
preparation reads for 29,080 bytes in every source-backed sample, matching
the change 0191 replay evidence, with zero bytes read from the eight
`Pictures/*` compressed ranges.

No allocation, RSS, physical-I/O, cold-cache, decompression, throughput,
scaling, edit/save, real-producer, broad ODF, or iWork claim is made. The
warm open-only p50/p99 acceptance is scoped to this corpus, host, build, and
selector.

## Verification

```text
cargo build --release --manifest-path tools/perf-baseline/Cargo.toml
sha256sum tools/perf-baseline/target/release/litchi-perf-baseline
taskset -c 2 <binary> --case odt_file_eager_open --warmup 30 --samples 500 --json a1.json
taskset -c 2 <binary> --case odt_file_source_open --warmup 30 --samples 500 --json b1.json
taskset -c 2 <binary> --case odt_file_source_open --warmup 30 --samples 500 --json b2.json
taskset -c 2 <binary> --case odt_file_eager_open --warmup 30 --samples 500 --json a2.json
```

The resulting binary hash equals the change 0191 binary hash, so no code
path differs between the two measurements.

Artifacts:

- [machine-readable summary](../results/odt-open-rerun-0192-summary.json)
- [artifact manifest](../results/odt-open-rerun-0192-manifest.json)
- compressed A1/B1/B2/A2 raw reports listed in the manifest
